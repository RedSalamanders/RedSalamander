#include "FileSystemMicrosoftDrive.h"

#include <algorithm>
#include <cstring>
#include <limits>

#include "Helpers.h"

HRESULT STDMETHODCALLTYPE FilesInformationMicrosoftDrive::QueryInterface(REFIID riid, void** ppvObject) noexcept
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

ULONG STDMETHODCALLTYPE FilesInformationMicrosoftDrive::AddRef() noexcept
{
    return _refCount.fetch_add(1, std::memory_order_relaxed) + 1;
}

ULONG STDMETHODCALLTYPE FilesInformationMicrosoftDrive::Release() noexcept
{
    const ULONG result = _refCount.fetch_sub(1, std::memory_order_acq_rel) - 1;
    if (result == 0)
    {
        delete this;
    }
    return result;
}

HRESULT STDMETHODCALLTYPE FilesInformationMicrosoftDrive::GetBuffer(FileInfo** ppFileInfo) noexcept
{
    return _packedBuffer.GetBuffer(ppFileInfo);
}

HRESULT STDMETHODCALLTYPE FilesInformationMicrosoftDrive::GetBufferSize(unsigned long* pSize) noexcept
{
    return _packedBuffer.GetBufferSize(pSize);
}

HRESULT STDMETHODCALLTYPE FilesInformationMicrosoftDrive::GetAllocatedSize(unsigned long* pSize) noexcept
{
    return _packedBuffer.GetAllocatedSize(pSize);
}

HRESULT STDMETHODCALLTYPE FilesInformationMicrosoftDrive::GetCount(unsigned long* pCount) noexcept
{
    return _packedBuffer.GetCount(pCount);
}

HRESULT STDMETHODCALLTYPE FilesInformationMicrosoftDrive::Get(unsigned long index, FileInfo** ppEntry) noexcept
{
    return _packedBuffer.Get(index, ppEntry);
}

HRESULT FilesInformationMicrosoftDrive::BuildFromEntries(std::vector<Entry> entries) noexcept
{
    std::sort(entries.begin(),
              entries.end(),
              [](const Entry& a, const Entry& b)
    {
        const bool aDir = (a.attributes & FILE_ATTRIBUTE_DIRECTORY) != 0;
        const bool bDir = (b.attributes & FILE_ATTRIBUTE_DIRECTORY) != 0;
        if (aDir != bDir)
        {
            return aDir;
        }

        const int cmp = OrdinalString::Compare(a.name, b.name, true);
        if (cmp != 0)
        {
            return cmp < 0;
        }

        return a.sizeBytes < b.sizeBytes;
    });

    return _packedBuffer.Build(entries,
                               [](const Entry& source, FileInfo& entry) noexcept
    {
        entry.FileAttributes = source.attributes;
        entry.FileIndex      = source.fileIndex;
        entry.EndOfFile      = static_cast<__int64>(source.sizeBytes);
        entry.AllocationSize = static_cast<__int64>(source.sizeBytes);
        entry.CreationTime   = source.creationTime;
        entry.LastAccessTime = source.lastAccessTime;
        entry.LastWriteTime  = source.lastWriteTime;
        entry.ChangeTime     = source.changeTime;
    });
}
