#include "FileSystemMtp.Internal.h"

#include <algorithm>
#include <chrono>
#include <cstring>
#include <cwctype>
#include <format>
#include <limits>
#include <thread>
#include <utility>

#include "Helpers.h"

namespace FileSystemMtpInternal
{
namespace
{
class MemoryBackendFileReader final : public IMtpBackendFileReader
{
public:
    MemoryBackendFileReader(std::vector<std::byte> bytes,
                            uint32_t readDelayMs,
                            std::shared_ptr<void> readContext,
                            MemoryBackendReadObserver observer) noexcept
        : _bytes(std::move(bytes)),
          _readDelayMs(readDelayMs),
          _readContext(std::move(readContext)),
          _observer(observer)
    {
    }

    MemoryBackendFileReader(const MemoryBackendFileReader&)            = delete;
    MemoryBackendFileReader(MemoryBackendFileReader&&)                 = delete;
    MemoryBackendFileReader& operator=(const MemoryBackendFileReader&) = delete;
    MemoryBackendFileReader& operator=(MemoryBackendFileReader&&)      = delete;

    HRESULT GetSize(uint64_t& sizeBytes) noexcept override
    {
        sizeBytes = static_cast<uint64_t>(_bytes.size());
        return S_OK;
    }

    HRESULT Seek(__int64 offset, unsigned long origin, uint64_t& newPosition) noexcept override
    {
        uint64_t base = 0;
        if (origin == FILE_CURRENT)
        {
            base = _position;
        }
        else if (origin == FILE_END)
        {
            base = static_cast<uint64_t>(_bytes.size());
        }
        else if (origin != FILE_BEGIN)
        {
            return E_INVALIDARG;
        }

        if (offset < 0)
        {
            const uint64_t delta = static_cast<uint64_t>(-(offset + 1)) + 1u;
            if (base < delta)
            {
                return HRESULT_FROM_WIN32(ERROR_NEGATIVE_SEEK);
            }
            _position = base - delta;
        }
        else
        {
            const uint64_t delta = static_cast<uint64_t>(offset);
            if (base > (std::numeric_limits<uint64_t>::max)() - delta)
            {
                return HRESULT_FROM_WIN32(ERROR_ARITHMETIC_OVERFLOW);
            }
            _position = base + delta;
        }

        newPosition = _position;
        return S_OK;
    }

    HRESULT Read(std::span<std::byte> buffer, unsigned long requestedBytes, unsigned long& bytesRead) noexcept override
    {
        bytesRead = 0;
        if (requestedBytes == 0u || _position >= static_cast<uint64_t>(_bytes.size()))
        {
            return S_OK;
        }
        if (buffer.empty())
        {
            return E_POINTER;
        }
        if (_readDelayMs != 0u)
        {
            std::this_thread::sleep_for(std::chrono::milliseconds(_readDelayMs));
        }

        const uint64_t available = static_cast<uint64_t>(_bytes.size()) - _position;
        const size_t requested   = std::min<size_t>(buffer.size(), requestedBytes);
        const size_t take        = static_cast<size_t>(std::min<uint64_t>(available, requested));
        if (take == 0u)
        {
            return S_OK;
        }

        std::memcpy(buffer.data(), _bytes.data() + _position, take);
        _position += static_cast<uint64_t>(take);
        bytesRead = static_cast<unsigned long>(take);
        if (_observer)
        {
            _observer(_readContext.get(), static_cast<uint64_t>(take));
        }
        Debug::Perf::EmitValue(L"mtp.transfer.read_bytes", static_cast<uint64_t>(take), S_OK);
        return S_OK;
    }

private:
    std::vector<std::byte> _bytes;
    uint32_t _readDelayMs = 0;
    uint64_t _position    = 0;
    std::shared_ptr<void> _readContext;
    MemoryBackendReadObserver _observer = nullptr;
};
} // namespace

std::shared_ptr<IMtpBackendFileReader> CreateMemoryBackendFileReader(std::vector<std::byte> bytes,
                                                                     uint32_t readDelayMs,
                                                                     std::shared_ptr<void> readContext,
                                                                     MemoryBackendReadObserver observer)
{
    return std::make_shared<MemoryBackendFileReader>(std::move(bytes), readDelayMs, std::move(readContext), observer);
}

[[nodiscard]] std::wstring NormalizeMtpPath(std::wstring_view rawPath) noexcept
{
    std::wstring path(rawPath);
    if (path.empty())
    {
        return L"/";
    }

    for (wchar_t& ch : path)
    {
        if (ch == L'\\')
        {
            ch = L'/';
        }
    }

    if (OrdinalString::StartsWithNoCase(path, L"mtp://"))
    {
        path.erase(0, 6);
        path.insert(path.begin(), L'/');
    }
    else if (OrdinalString::StartsWithNoCase(path, L"mtp:/"))
    {
        path.erase(0, 4);
    }

    while (path.size() >= 2u && path[0] == L'/' && path[1] == L'/')
    {
        path.erase(path.begin());
    }

    if (path.empty() || path.front() != L'/')
    {
        path.insert(path.begin(), L'/');
    }

    std::wstring collapsed;
    collapsed.reserve(path.size());
    bool previousSlash = false;
    for (const wchar_t ch : path)
    {
        const bool slash = ch == L'/';
        if (slash && previousSlash)
        {
            continue;
        }

        collapsed.push_back(ch);
        previousSlash = slash;
    }

    while (collapsed.size() > 1u && collapsed.back() == L'/')
    {
        collapsed.pop_back();
    }

    return collapsed.empty() ? std::wstring(L"/") : collapsed;
}

[[nodiscard]] std::vector<std::wstring_view> SplitPathSegments(std::wstring_view path) noexcept
{
    std::vector<std::wstring_view> segments;
    while (! path.empty() && path.front() == L'/')
    {
        path.remove_prefix(1);
    }

    while (! path.empty())
    {
        const size_t slash           = path.find(L'/');
        const std::wstring_view part = (slash == std::wstring_view::npos) ? path : path.substr(0, slash);
        if (! part.empty())
        {
            segments.push_back(part);
        }
        if (slash == std::wstring_view::npos)
        {
            break;
        }
        path.remove_prefix(slash + 1u);
        while (! path.empty() && path.front() == L'/')
        {
            path.remove_prefix(1);
        }
    }

    return segments;
}

[[nodiscard]] std::wstring ParentPath(std::wstring_view path) noexcept
{
    const std::wstring normalized = NormalizeMtpPath(path);
    if (normalized == L"/")
    {
        return L"/";
    }

    const size_t slash = normalized.find_last_of(L'/');
    if (slash == std::wstring::npos || slash == 0)
    {
        return L"/";
    }

    return normalized.substr(0, slash);
}

[[nodiscard]] std::wstring LeafName(std::wstring_view path) noexcept
{
    const std::wstring normalized = NormalizeMtpPath(path);
    if (normalized == L"/")
    {
        return {};
    }

    const size_t slash = normalized.find_last_of(L'/');
    if (slash == std::wstring::npos)
    {
        return normalized;
    }

    return normalized.substr(slash + 1u);
}

[[nodiscard]] std::wstring JoinPath(std::wstring_view parent, std::wstring_view leaf) noexcept
{
    std::wstring normalizedParent = NormalizeMtpPath(parent);
    if (leaf.empty())
    {
        return normalizedParent;
    }

    if (normalizedParent == L"/")
    {
        std::wstring result;
        result.reserve(leaf.size() + 1u);
        result.push_back(L'/');
        result.append(leaf);
        return NormalizeMtpPath(result);
    }

    std::wstring result(normalizedParent);
    result.push_back(L'/');
    result.append(leaf);
    return NormalizeMtpPath(result);
}

[[nodiscard]] std::string Utf8FromUtf16(std::wstring_view text) noexcept
{
    return Common::Strings::Utf8FromUtf16StrictOrEmpty(text);
}

[[nodiscard]] std::wstring Utf16FromUtf8(std::string_view text) noexcept
{
    return Common::Strings::Utf16FromUtf8StrictOrEmpty(text);
}

[[nodiscard]] std::uint64_t StableMtpIdentityHash(std::wstring_view value) noexcept
{
    std::uint64_t hash = 1469598103934665603ull;
    for (const wchar_t ch : value)
    {
        const auto lower = static_cast<std::uint64_t>(::towlower(static_cast<wint_t>(ch)));
        hash ^= lower & 0xFFu;
        hash *= 1099511628211ull;
        hash ^= (lower >> 8u) & 0xFFu;
        hash *= 1099511628211ull;
    }
    return hash;
}

[[nodiscard]] std::wstring FormatMtpIdentityHash(std::wstring_view value)
{
    return std::format(L"{:016X}", StableMtpIdentityHash(value));
}

[[nodiscard]] std::wstring SanitizeMtpPathComponent(std::wstring value)
{
    for (auto& ch : value)
    {
        if (ch == L'/' || ch == L'\\' || ch == L'\0')
        {
            ch = L'_';
        }
    }
    return value;
}

[[nodiscard]] std::wstring MtpDeviceIdentitySuffix(std::wstring_view pnpId)
{
    return std::format(L"[devid:{}]", FormatMtpIdentityHash(pnpId));
}

[[nodiscard]] std::wstring MtpPersistentObjectIdentitySuffix(std::wstring_view persistentId)
{
    return std::format(L" [puid:{}]", FormatMtpIdentityHash(persistentId));
}

[[nodiscard]] std::wstring MtpObjectIdentitySuffix(std::wstring_view objectId)
{
    return std::format(L" [oid:{}]", FormatMtpIdentityHash(objectId));
}

[[nodiscard]] std::wstring MtpDuplicateObjectSuffix(const MtpItem& item)
{
    return item.persistentId.empty() ? MtpObjectIdentitySuffix(item.objectId) : MtpPersistentObjectIdentitySuffix(item.persistentId);
}

[[nodiscard]] std::string JsonEscapeUtf8(std::string_view text)
{
    std::string out;
    out.reserve(text.size() + 8u);
    for (const char value : text)
    {
        const unsigned char ch = static_cast<unsigned char>(value);
        switch (ch)
        {
            case '\\': out.append("\\\\"); break;
            case '"': out.append("\\\""); break;
            case '\b': out.append("\\b"); break;
            case '\f': out.append("\\f"); break;
            case '\n': out.append("\\n"); break;
            case '\r': out.append("\\r"); break;
            case '\t': out.append("\\t"); break;
            default:
            {
                if (ch < 0x20u)
                {
                    constexpr char kHex[] = "0123456789ABCDEF";
                    out.append("\\u00");
                    out.push_back(kHex[(ch >> 4u) & 0xFu]);
                    out.push_back(kHex[ch & 0xFu]);
                }
                else
                {
                    out.push_back(value);
                }
                break;
            }
        }
    }
    return out;
}

[[nodiscard]] bool EqualsPathComponent(std::wstring_view a, std::wstring_view b) noexcept
{
    return a == b;
}

[[nodiscard]] __int64 SystemTimeToFileTime64(const SYSTEMTIME& value) noexcept
{
    FILETIME ft{};
    if (SystemTimeToFileTime(&value, &ft) == 0)
    {
        return 0;
    }

    ULARGE_INTEGER u{};
    u.LowPart  = ft.dwLowDateTime;
    u.HighPart = ft.dwHighDateTime;
    if (u.QuadPart > static_cast<ULONGLONG>((std::numeric_limits<__int64>::max)()))
    {
        return 0;
    }

    return static_cast<__int64>(u.QuadPart);
}

[[nodiscard]] __int64 NowFileTime64() noexcept
{
    FILETIME ft{};
    GetSystemTimeAsFileTime(&ft);
    ULARGE_INTEGER u{};
    u.LowPart  = ft.dwLowDateTime;
    u.HighPart = ft.dwHighDateTime;
    return static_cast<__int64>(u.QuadPart);
}
} // namespace FileSystemMtpInternal
