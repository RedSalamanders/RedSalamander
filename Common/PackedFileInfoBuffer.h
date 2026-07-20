#pragma once

#include <cstddef>
#include <cstring>
#include <limits>
#include <string_view>
#include <utility>
#include <vector>

#include "PlugInterfaces/FileSystem.h"

namespace Common::Plugins
{
class PackedFileInfoBuffer final
{
public:
    PackedFileInfoBuffer()                                       = default;
    PackedFileInfoBuffer(const PackedFileInfoBuffer&)            = delete;
    PackedFileInfoBuffer(PackedFileInfoBuffer&&)                 = delete;
    PackedFileInfoBuffer& operator=(const PackedFileInfoBuffer&) = delete;
    PackedFileInfoBuffer& operator=(PackedFileInfoBuffer&&)      = delete;
    ~PackedFileInfoBuffer()                                      = default;

    HRESULT GetBuffer(FileInfo** fileInfo) noexcept
    {
        if (! fileInfo)
        {
            return E_POINTER;
        }

        *fileInfo = _buffer.empty() ? nullptr : reinterpret_cast<FileInfo*>(_buffer.data());
        return S_OK;
    }

    HRESULT GetBufferSize(unsigned long* size) const noexcept
    {
        if (! size)
        {
            return E_POINTER;
        }

        *size = _usedBytes;
        return S_OK;
    }

    HRESULT GetAllocatedSize(unsigned long* size) const noexcept
    {
        if (! size)
        {
            return E_POINTER;
        }

        *size = static_cast<unsigned long>(_buffer.size());
        return S_OK;
    }

    HRESULT GetCount(unsigned long* count) const noexcept
    {
        if (! count)
        {
            return E_POINTER;
        }

        *count = _count;
        return S_OK;
    }

    HRESULT Get(unsigned long index, FileInfo** entry) noexcept
    {
        if (! entry)
        {
            return E_POINTER;
        }

        *entry = nullptr;
        if (index >= _count)
        {
            return HRESULT_FROM_WIN32(ERROR_NO_MORE_FILES);
        }

        return LocateEntry(index, entry);
    }

    template <typename TEntry, typename TPopulate> HRESULT Build(const std::vector<TEntry>& entries, TPopulate&& populate) noexcept
    {
        Reset();
        if (entries.empty())
        {
            return S_OK;
        }
        if (entries.size() > static_cast<size_t>((std::numeric_limits<unsigned long>::max)()))
        {
            return HRESULT_FROM_WIN32(ERROR_ARITHMETIC_OVERFLOW);
        }

        size_t totalBytes = 0;
        for (const TEntry& source : entries)
        {
            size_t entrySize     = 0;
            const HRESULT sizeHr = TryComputeEntrySize(std::wstring_view(source.name), entrySize);
            if (FAILED(sizeHr) || entrySize > static_cast<size_t>((std::numeric_limits<unsigned long>::max)()) ||
                totalBytes > static_cast<size_t>((std::numeric_limits<unsigned long>::max)()) - entrySize)
            {
                return FAILED(sizeHr) ? sizeHr : HRESULT_FROM_WIN32(ERROR_ARITHMETIC_OVERFLOW);
            }
            totalBytes += entrySize;
        }

        std::vector<std::byte> buffer(totalBytes, std::byte{0});
        size_t offset       = 0;
        FileInfo* previous  = nullptr;
        size_t previousSize = 0;
        for (const TEntry& source : entries)
        {
            size_t entrySize     = 0;
            const HRESULT sizeHr = TryComputeEntrySize(std::wstring_view(source.name), entrySize);
            if (FAILED(sizeHr) || entrySize > buffer.size() - offset)
            {
                return FAILED(sizeHr) ? sizeHr : HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
            }

            auto* entry = reinterpret_cast<FileInfo*>(buffer.data() + offset);
            populate(source, *entry);

            const size_t nameBytes = source.name.size() * sizeof(wchar_t);
            entry->FileNameSize    = static_cast<unsigned long>(nameBytes);
            if (nameBytes != 0)
            {
                std::memcpy(entry->FileName, source.name.data(), nameBytes);
            }
            entry->FileName[source.name.size()] = L'\0';

            if (previous)
            {
                previous->NextEntryOffset = static_cast<unsigned long>(previousSize);
            }
            previous     = entry;
            previousSize = entrySize;
            offset += entrySize;
        }

        _buffer    = std::move(buffer);
        _count     = static_cast<unsigned long>(entries.size());
        _usedBytes = static_cast<unsigned long>(totalBytes);
        return S_OK;
    }

private:
    static constexpr size_t kEntryAlignment = alignof(FileInfo);

    static HRESULT TryAlignUp(size_t value, size_t& aligned) noexcept
    {
        static_assert((kEntryAlignment & (kEntryAlignment - 1u)) == 0u);
        constexpr size_t mask = kEntryAlignment - 1u;
        if (value > (std::numeric_limits<size_t>::max)() - mask)
        {
            return HRESULT_FROM_WIN32(ERROR_ARITHMETIC_OVERFLOW);
        }

        aligned = (value + mask) & ~mask;
        return S_OK;
    }

    static HRESULT TryComputeEntrySize(std::wstring_view name, size_t& entrySize) noexcept
    {
        constexpr size_t baseSize = offsetof(FileInfo, FileName);
        if (name.size() > static_cast<size_t>((std::numeric_limits<unsigned long>::max)()) / sizeof(wchar_t))
        {
            return HRESULT_FROM_WIN32(ERROR_ARITHMETIC_OVERFLOW);
        }

        const size_t nameBytes = name.size() * sizeof(wchar_t);
        if (nameBytes > (std::numeric_limits<size_t>::max)() - baseSize - sizeof(wchar_t))
        {
            return HRESULT_FROM_WIN32(ERROR_ARITHMETIC_OVERFLOW);
        }

        return TryAlignUp(baseSize + nameBytes + sizeof(wchar_t), entrySize);
    }

    HRESULT LocateEntry(unsigned long index, FileInfo** result) noexcept
    {
        size_t offset = 0;
        for (unsigned long current = 0; current < _count; ++current)
        {
            if (offset > _buffer.size() || _buffer.size() - offset < sizeof(FileInfo))
            {
                return HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
            }

            auto* entry = reinterpret_cast<FileInfo*>(_buffer.data() + offset);
            if ((entry->FileNameSize % sizeof(wchar_t)) != 0u)
            {
                return HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
            }

            size_t entrySize = 0;
            const HRESULT sizeHr =
                TryComputeEntrySize(std::wstring_view(entry->FileName, static_cast<size_t>(entry->FileNameSize) / sizeof(wchar_t)), entrySize);
            if (FAILED(sizeHr) || entrySize > _buffer.size() - offset)
            {
                return HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
            }

            if (current == index)
            {
                *result = entry;
                return S_OK;
            }

            const size_t advance = static_cast<size_t>(entry->NextEntryOffset);
            if (advance < entrySize || (advance % kEntryAlignment) != 0u || advance > _buffer.size() - offset)
            {
                return HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
            }
            offset += advance;
        }

        return HRESULT_FROM_WIN32(ERROR_NO_MORE_FILES);
    }

    void Reset() noexcept
    {
        _buffer.clear();
        _count     = 0;
        _usedBytes = 0;
    }

    std::vector<std::byte> _buffer;
    unsigned long _count     = 0;
    unsigned long _usedBytes = 0;
};
} // namespace Common::Plugins
