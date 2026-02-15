#include "FileSystemS3.Internal.h"

// FilesInformationS3

HRESULT STDMETHODCALLTYPE FilesInformationS3::QueryInterface(REFIID riid, void** ppvObject) noexcept
{
    if (ppvObject == nullptr)
    {
        return E_POINTER;
    }

    if (riid == __uuidof(IUnknown) || riid == __uuidof(IFilesInformation))
    {
        *ppvObject = static_cast<IFilesInformation*>(this);
        AddRef();
        return S_OK;
    }

    *ppvObject = nullptr;
    return E_NOINTERFACE;
}

ULONG STDMETHODCALLTYPE FilesInformationS3::AddRef() noexcept
{
    return _refCount.fetch_add(1, std::memory_order_relaxed) + 1;
}

ULONG STDMETHODCALLTYPE FilesInformationS3::Release() noexcept
{
    const ULONG result = _refCount.fetch_sub(1, std::memory_order_acq_rel) - 1;
    if (result == 0)
    {
        delete this;
    }
    return result;
}

HRESULT STDMETHODCALLTYPE FilesInformationS3::GetBuffer(FileInfo** ppFileInfo) noexcept
{
    if (ppFileInfo == nullptr)
    {
        return E_POINTER;
    }

    *ppFileInfo = nullptr;

    if (_usedBytes == 0 || _buffer.empty())
    {
        return S_OK;
    }

    *ppFileInfo = reinterpret_cast<FileInfo*>(_buffer.data());
    return S_OK;
}

HRESULT STDMETHODCALLTYPE FilesInformationS3::GetBufferSize(unsigned long* pSize) noexcept
{
    if (pSize == nullptr)
    {
        return E_POINTER;
    }

    *pSize = _usedBytes;
    return S_OK;
}

HRESULT STDMETHODCALLTYPE FilesInformationS3::GetAllocatedSize(unsigned long* pSize) noexcept
{
    if (pSize == nullptr)
    {
        return E_POINTER;
    }

    if (_buffer.size() > static_cast<size_t>((std::numeric_limits<unsigned long>::max)()))
    {
        return HRESULT_FROM_WIN32(ERROR_ARITHMETIC_OVERFLOW);
    }

    *pSize = static_cast<unsigned long>(_buffer.size());
    return S_OK;
}

HRESULT STDMETHODCALLTYPE FilesInformationS3::GetCount(unsigned long* pCount) noexcept
{
    if (pCount == nullptr)
    {
        return E_POINTER;
    }

    *pCount = _count;
    return S_OK;
}

HRESULT STDMETHODCALLTYPE FilesInformationS3::Get(unsigned long index, FileInfo** ppEntry) noexcept
{
    if (ppEntry == nullptr)
    {
        return E_POINTER;
    }

    *ppEntry = nullptr;

    if (index >= _count)
    {
        return HRESULT_FROM_WIN32(ERROR_NO_MORE_FILES);
    }

    return LocateEntry(index, ppEntry);
}

size_t FilesInformationS3::AlignUp(size_t value, size_t alignment) noexcept
{
    const size_t mask = alignment - 1u;
    return (value + mask) & ~mask;
}

size_t FilesInformationS3::ComputeEntrySizeBytes(std::wstring_view name) noexcept
{
    const size_t baseSize = offsetof(FileInfo, FileName);
    const size_t nameSize = name.size() * sizeof(wchar_t);
    return AlignUp(baseSize + nameSize + sizeof(wchar_t), sizeof(unsigned long));
}

HRESULT FilesInformationS3::BuildFromEntries(std::vector<Entry> entries) noexcept
{
    _buffer.clear();
    _count     = 0;
    _usedBytes = 0;

    if (entries.empty())
    {
        return S_OK;
    }

    std::sort(entries.begin(),
              entries.end(),
              [](const Entry& a, const Entry& b)
              {
                  const int cmp = OrdinalString::Compare(a.name, b.name, true);
                  if (cmp != 0)
                  {
                      return cmp < 0;
                  }

                  const bool aDir = (a.attributes & FILE_ATTRIBUTE_DIRECTORY) != 0;
                  const bool bDir = (b.attributes & FILE_ATTRIBUTE_DIRECTORY) != 0;
                  if (aDir != bDir)
                  {
                      return aDir;
                  }

                  return a.sizeBytes < b.sizeBytes;
              });

    size_t totalBytes = 0;
    for (const auto& entry : entries)
    {
        totalBytes += ComputeEntrySizeBytes(entry.name);
        if (totalBytes > static_cast<size_t>((std::numeric_limits<unsigned long>::max)()))
        {
            return HRESULT_FROM_WIN32(ERROR_ARITHMETIC_OVERFLOW);
        }
    }

    _buffer.resize(totalBytes, std::byte{0});

    std::byte* base     = _buffer.data();
    size_t offset       = 0;
    FileInfo* previous  = nullptr;
    size_t previousSize = 0;

    for (const auto& source : entries)
    {
        const size_t entrySize = ComputeEntrySizeBytes(source.name);
        if (offset + entrySize > _buffer.size())
        {
            return E_FAIL;
        }

        auto* entry = reinterpret_cast<FileInfo*>(base + offset);
        std::memset(entry, 0, entrySize);

        const size_t nameBytes = source.name.size() * sizeof(wchar_t);
        if (nameBytes > static_cast<size_t>((std::numeric_limits<unsigned long>::max)()))
        {
            return HRESULT_FROM_WIN32(ERROR_ARITHMETIC_OVERFLOW);
        }

        entry->FileAttributes = source.attributes;
        entry->FileIndex      = source.fileIndex;
        entry->EndOfFile      = static_cast<__int64>(source.sizeBytes);
        entry->AllocationSize = static_cast<__int64>(source.sizeBytes);

        entry->CreationTime   = source.creationTime;
        entry->LastAccessTime = source.lastAccessTime;
        entry->LastWriteTime  = source.lastWriteTime;
        entry->ChangeTime     = source.changeTime;

        entry->FileNameSize = static_cast<unsigned long>(nameBytes);
        if (! source.name.empty())
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
        ++_count;
    }

    _usedBytes = static_cast<unsigned long>(_buffer.size());
    return S_OK;
}

HRESULT FilesInformationS3::LocateEntry(unsigned long index, FileInfo** ppEntry) const noexcept
{
    const std::byte* base      = _buffer.data();
    size_t offset              = 0;
    unsigned long currentIndex = 0;

    while (offset < _usedBytes && offset + sizeof(FileInfo) <= _buffer.size())
    {
        auto* entry = reinterpret_cast<const FileInfo*>(base + offset);
        if (currentIndex == index)
        {
            *ppEntry = const_cast<FileInfo*>(entry);
            return S_OK;
        }

        const size_t advance = (entry->NextEntryOffset != 0)
                                   ? static_cast<size_t>(entry->NextEntryOffset)
                                   : ComputeEntrySizeBytes(std::wstring_view(entry->FileName, static_cast<size_t>(entry->FileNameSize) / sizeof(wchar_t)));
        if (advance == 0)
        {
            break;
        }

        offset += advance;
        ++currentIndex;
    }

    return HRESULT_FROM_WIN32(ERROR_NO_MORE_FILES);
}
