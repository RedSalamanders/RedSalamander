#include "SessionState.h"

#include "Framework.h"

#include <atomic>
#include <cstddef>
#include <limits>
#include <mutex>
#include <system_error>

#pragma warning(push)
// WIL: C4625 (copy ctor deleted), C4626 (copy assign deleted), C5026 (move ctor deleted), C5027 (move assign deleted)
#pragma warning(disable : 4625 4626 5026 5027)
#include <wil/resource.h>
#pragma warning(pop)

#include "AppDataPaths.h"
#include "Helpers.h"

namespace SessionState
{
namespace
{
constexpr wchar_t kCompanyDirName[]  = L"RedSalamander";
constexpr wchar_t kSessionFileName[] = L"session_state.txt";

std::mutex g_mutex;
std::atomic<OperationKind> g_lastOperation{OperationKind::Unknown};
std::vector<std::wstring> g_lastActiveFileSystemPluginIds;

[[nodiscard]] bool ContainsNoCase(const std::vector<std::wstring>& ids, std::wstring_view id) noexcept
{
    if (id.empty())
    {
        return false;
    }

    for (const std::wstring& existing : ids)
    {
        if (OrdinalString::EqualsNoCase(existing, id))
        {
            return true;
        }
    }
    return false;
}

[[nodiscard]] bool EqualNoCase(const std::vector<std::wstring>& a, const std::vector<std::wstring>& b) noexcept
{
    if (a.size() != b.size())
    {
        return false;
    }

    for (size_t i = 0; i < a.size(); ++i)
    {
        if (! OrdinalString::EqualsNoCase(a[i], b[i]))
        {
            return false;
        }
    }

    return true;
}

[[nodiscard]] std::vector<std::wstring> NormalizePluginIds(std::initializer_list<std::wstring_view> pluginIds)
{
    std::vector<std::wstring> normalized;
    normalized.reserve(pluginIds.size());
    for (const std::wstring_view id : pluginIds)
    {
        if (id.empty())
        {
            continue;
        }

        if (! ContainsNoCase(normalized, id))
        {
            normalized.emplace_back(id);
        }
    }

    return normalized;
}

[[nodiscard]] std::filesystem::path GetSessionStateDirectory() noexcept
{
    const std::filesystem::path base = AppDataPaths::GetLocalAppDataPath();
    if (base.empty())
    {
        return {};
    }

    return base / kCompanyDirName;
}

[[nodiscard]] HRESULT EnsureDirectoryExists(const std::filesystem::path& dir) noexcept
{
    if (dir.empty())
    {
        return E_INVALIDARG;
    }

    std::error_code ec;
    if (std::filesystem::exists(dir, ec))
    {
        return S_OK;
    }

    std::filesystem::create_directories(dir, ec);
    return ec ? HRESULT_FROM_WIN32(static_cast<DWORD>(ec.value())) : S_OK;
}

[[nodiscard]] std::wstring_view OperationToText(OperationKind op) noexcept
{
    switch (op)
    {
        case OperationKind::Browse: return L"browse";
        case OperationKind::Copy: return L"copy";
        case OperationKind::Compare: return L"compare";
        case OperationKind::Unknown: break;
    }
    return L"unknown";
}

[[nodiscard]] OperationKind ParseOperation(std::wstring_view text) noexcept
{
    if (text == L"browse")
    {
        return OperationKind::Browse;
    }
    if (text == L"copy")
    {
        return OperationKind::Copy;
    }
    if (text == L"compare")
    {
        return OperationKind::Compare;
    }
    return OperationKind::Unknown;
}

void WriteMarkerFileLocked(const std::vector<std::wstring>& pluginIds, OperationKind op) noexcept
{
    const std::filesystem::path dir  = GetSessionStateDirectory();
    const std::filesystem::path path = dir.empty() ? std::filesystem::path{} : (dir / kSessionFileName);
    if (path.empty())
    {
        return;
    }

    if (FAILED(EnsureDirectoryExists(dir)))
    {
        return;
    }

    wil::unique_handle file(CreateFileW(path.c_str(), GENERIC_WRITE, FILE_SHARE_READ, nullptr, CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, nullptr));
    if (! file)
    {
        return;
    }

    const wchar_t bom = 0xFEFF;
    DWORD written     = 0;
    if (! WriteFile(file.get(), &bom, sizeof(bom), &written, nullptr))
    {
        return;
    }

    std::wstring content;
    size_t reserveHint = 64u;
    for (const std::wstring& id : pluginIds)
    {
        reserveHint += id.size() + 16u;
    }
    content.reserve(reserveHint);

    size_t writtenPlugins = 0;
    for (const std::wstring& id : pluginIds)
    {
        if (id.empty())
        {
            continue;
        }

        ++writtenPlugins;
        if (writtenPlugins == 1)
        {
            content.append(L"fsPlugin=");
        }
        else
        {
            content.append(L"fsPlugin");
            content.append(std::format(L"{:d}", writtenPlugins));
            content.append(L"=");
        }
        content.append(id);
        content.append(L"\r\n");
    }

    content.append(L"op=");
    content.append(OperationToText(op));
    content.append(L"\r\n");

    const DWORD bytes = static_cast<DWORD>(content.size() * sizeof(wchar_t));
    if (bytes > 0)
    {
        static_cast<void>(WriteFile(file.get(), content.data(), bytes, &written, nullptr));
    }

    static_cast<void>(FlushFileBuffers(file.get()));
}

[[nodiscard]] std::optional<std::wstring> ReadUtf16File(const std::filesystem::path& path) noexcept
{
    if (path.empty())
    {
        return std::nullopt;
    }

    wil::unique_handle file(CreateFileW(
        path.c_str(), GENERIC_READ, FILE_SHARE_READ | FILE_SHARE_WRITE | FILE_SHARE_DELETE, nullptr, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, nullptr));
    if (! file)
    {
        return std::nullopt;
    }

    LARGE_INTEGER size{};
    if (GetFileSizeEx(file.get(), &size) == 0)
    {
        return std::nullopt;
    }

    if (size.QuadPart <= 0 || (size.QuadPart % 2) != 0)
    {
        return std::nullopt;
    }

    const uint64_t bytes = static_cast<uint64_t>(size.QuadPart);
    if (bytes > (std::numeric_limits<size_t>::max)())
    {
        return std::nullopt;
    }

    std::wstring buffer;
    buffer.resize(static_cast<size_t>(bytes / 2u), L'\0');

    DWORD readBytes = 0;
    if (ReadFile(file.get(), buffer.data(), static_cast<DWORD>(bytes), &readBytes, nullptr) == 0)
    {
        return std::nullopt;
    }
    buffer.resize(static_cast<size_t>(readBytes / 2u));

    if (! buffer.empty() && buffer.front() == 0xFEFF)
    {
        buffer.erase(buffer.begin());
    }

    return buffer;
}

[[nodiscard]] std::wstring_view Trim(std::wstring_view text) noexcept
{
    while (! text.empty())
    {
        const wchar_t ch = text.front();
        if (ch != L' ' && ch != L'\t' && ch != L'\r' && ch != L'\n')
        {
            break;
        }
        text.remove_prefix(1);
    }
    while (! text.empty())
    {
        const wchar_t ch = text.back();
        if (ch != L' ' && ch != L'\t' && ch != L'\r' && ch != L'\n')
        {
            break;
        }
        text.remove_suffix(1);
    }
    return text;
}
} // namespace

std::filesystem::path GetSessionStatePath() noexcept
{
    const std::filesystem::path dir = GetSessionStateDirectory();
    return dir.empty() ? std::filesystem::path{} : (dir / kSessionFileName);
}

void Clear() noexcept
{
    const std::filesystem::path path = GetSessionStatePath();
    if (path.empty())
    {
        return;
    }

    std::error_code ec;
    static_cast<void>(std::filesystem::remove(path, ec));

    std::scoped_lock lock(g_mutex);
    g_lastActiveFileSystemPluginIds.clear();
    g_lastOperation.store(OperationKind::Unknown, std::memory_order_release);
}

void UpdateActiveFileSystemPluginIdsAndOperation(std::initializer_list<std::wstring_view> pluginIds, OperationKind operation) noexcept
{
    std::vector<std::wstring> normalized = NormalizePluginIds(pluginIds);

    std::scoped_lock lock(g_mutex);
    const bool pluginsChanged    = ! EqualNoCase(normalized, g_lastActiveFileSystemPluginIds);
    const OperationKind previous = g_lastOperation.load(std::memory_order_acquire);
    const bool operationChanged  = previous != operation;
    if (! pluginsChanged && ! operationChanged)
    {
        return;
    }

    if (pluginsChanged)
    {
        g_lastActiveFileSystemPluginIds = std::move(normalized);
    }
    g_lastOperation.store(operation, std::memory_order_release);
    WriteMarkerFileLocked(g_lastActiveFileSystemPluginIds, operation);
}

std::optional<State> TryRead() noexcept
{
    const std::filesystem::path path = GetSessionStatePath();
    if (path.empty())
    {
        return std::nullopt;
    }

    const auto textOpt = ReadUtf16File(path);
    if (! textOpt.has_value())
    {
        return std::nullopt;
    }

    State state{};

    std::wstring_view text(textOpt.value());
    size_t start = 0;
    while (start < text.size())
    {
        const size_t end       = text.find_first_of(L"\r\n", start);
        const size_t lineEnd   = (end == std::wstring_view::npos) ? text.size() : end;
        std::wstring_view line = text.substr(start, lineEnd - start);
        line                   = Trim(line);
        if (! line.empty())
        {
            const size_t eq = line.find(L'=');
            if (eq != std::wstring_view::npos)
            {
                const std::wstring_view key   = Trim(line.substr(0, eq));
                const std::wstring_view value = Trim(line.substr(eq + 1));
                if (key.rfind(L"fsPlugin", 0) == 0)
                {
                    if (! value.empty() && ! ContainsNoCase(state.activeFileSystemPluginIds, value))
                    {
                        state.activeFileSystemPluginIds.emplace_back(value);
                    }
                }
                else if (key == L"op")
                {
                    state.lastOperation = ParseOperation(value);
                }
            }
        }

        if (end == std::wstring_view::npos)
        {
            break;
        }

        // Skip \r\n or single newline.
        start = end + 1;
        if (start < text.size() && text[end] == L'\r' && text[start] == L'\n')
        {
            ++start;
        }
    }

    return state;
}
} // namespace SessionState
