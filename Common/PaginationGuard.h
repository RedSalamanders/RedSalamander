#pragma once

#include <cstddef>
#include <cstdint>
#include <limits>
#include <string>
#include <string_view>
#include <unordered_set>

#ifndef WIN32_LEAN_AND_MEAN
#define WIN32_LEAN_AND_MEAN
#endif
#ifndef NOMINMAX
#define NOMINMAX
#endif
#include <Windows.h>

namespace Common::Paging
{
using CancellationProbe = HRESULT (*)(void* cookie) noexcept;

struct Limits
{
    size_t maxPages                     = 4'096u;
    size_t maxItems                     = 4'000'000u;
    size_t maxBytes                     = 512u * 1024u * 1024u;
    size_t maxTokenChars                = 64u * 1024u;
    uint64_t deadlineTickMs             = 0u;
    CancellationProbe cancellationProbe = nullptr;
    void* cancellationCookie            = nullptr;
};

[[nodiscard]] inline uint64_t DeadlineFromNow(uint64_t nowTickMs, uint64_t durationMs) noexcept
{
    return durationMs > (std::numeric_limits<uint64_t>::max)() - nowTickMs ? (std::numeric_limits<uint64_t>::max)() : nowTickMs + durationMs;
}

template <typename CharT> class ContinuationGuard final
{
public:
    explicit ContinuationGuard(Limits limits) : _limits(limits)
    {
    }

    [[nodiscard]] HRESULT BeginFirstPage(uint64_t nowTickMs) noexcept
    {
        return BeginFirstPage({}, nowTickMs);
    }

    [[nodiscard]] HRESULT BeginFirstPage(std::basic_string_view<CharT> firstPageIdentity, uint64_t nowTickMs) noexcept
    {
        if (_started || firstPageIdentity.size() > _limits.maxTokenChars)
        {
            return HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
        }
        if (! firstPageIdentity.empty())
        {
            _seenTokens.emplace(firstPageIdentity);
        }
        _started = true;
        return BeginPage(nowTickMs);
    }

    [[nodiscard]] HRESULT BeginContinuation(std::basic_string_view<CharT> token, uint64_t nowTickMs) noexcept
    {
        if (! _started || _pageOpen || token.empty() || token.size() > _limits.maxTokenChars)
        {
            return HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
        }
        if (! _seenTokens.emplace(token).second)
        {
            return HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
        }
        return BeginPage(nowTickMs);
    }

    [[nodiscard]] HRESULT CompletePage(
        size_t pageItems, size_t pageBytes, bool serverSaysMore, std::basic_string_view<CharT> nextToken, uint64_t nowTickMs) noexcept
    {
        if (! _pageOpen)
        {
            return HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
        }
        _pageOpen = false;

        HRESULT hr = CheckBoundary(nowTickMs);
        if (FAILED(hr))
        {
            return hr;
        }
        if (pageItems > _limits.maxItems - _items || pageBytes > _limits.maxBytes - _bytes)
        {
            return HRESULT_FROM_WIN32(ERROR_FILE_TOO_LARGE);
        }
        _items += pageItems;
        _bytes += pageBytes;

        if (serverSaysMore && (nextToken.empty() || nextToken.size() > _limits.maxTokenChars))
        {
            return HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
        }
        return S_OK;
    }

    [[nodiscard]] size_t PageCount() const noexcept
    {
        return _pages;
    }
    [[nodiscard]] size_t ItemCount() const noexcept
    {
        return _items;
    }
    [[nodiscard]] size_t ByteCount() const noexcept
    {
        return _bytes;
    }

private:
    [[nodiscard]] HRESULT BeginPage(uint64_t nowTickMs) noexcept
    {
        if (_pageOpen || _limits.maxPages == 0u || _limits.maxItems == 0u || _limits.maxBytes == 0u || _limits.maxTokenChars == 0u)
        {
            return E_INVALIDARG;
        }
        HRESULT hr = CheckBoundary(nowTickMs);
        if (FAILED(hr))
        {
            return hr;
        }
        if (_pages >= _limits.maxPages)
        {
            return HRESULT_FROM_WIN32(ERROR_FILE_TOO_LARGE);
        }
        ++_pages;
        _pageOpen = true;
        return S_OK;
    }

    [[nodiscard]] HRESULT CheckBoundary(uint64_t nowTickMs) const noexcept
    {
        if (_limits.cancellationProbe)
        {
            const HRESULT cancelHr = _limits.cancellationProbe(_limits.cancellationCookie);
            if (FAILED(cancelHr))
            {
                return cancelHr;
            }
        }
        if (_limits.deadlineTickMs != 0u && nowTickMs >= _limits.deadlineTickMs)
        {
            return HRESULT_FROM_WIN32(ERROR_TIMEOUT);
        }
        return S_OK;
    }

    Limits _limits;
    std::unordered_set<std::basic_string<CharT>> _seenTokens;
    size_t _pages  = 0u;
    size_t _items  = 0u;
    size_t _bytes  = 0u;
    bool _started  = false;
    bool _pageOpen = false;
};

using Utf8ContinuationGuard = ContinuationGuard<char>;
using WideContinuationGuard = ContinuationGuard<wchar_t>;
} // namespace Common::Paging
