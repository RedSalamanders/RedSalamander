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
    return _packedBuffer.GetBuffer(ppFileInfo);
}

HRESULT STDMETHODCALLTYPE FilesInformationS3::GetBufferSize(unsigned long* pSize) noexcept
{
    return _packedBuffer.GetBufferSize(pSize);
}

HRESULT STDMETHODCALLTYPE FilesInformationS3::GetAllocatedSize(unsigned long* pSize) noexcept
{
    return _packedBuffer.GetAllocatedSize(pSize);
}

HRESULT STDMETHODCALLTYPE FilesInformationS3::GetCount(unsigned long* pCount) noexcept
{
    return _packedBuffer.GetCount(pCount);
}

HRESULT STDMETHODCALLTYPE FilesInformationS3::Get(unsigned long index, FileInfo** ppEntry) noexcept
{
    return _packedBuffer.Get(index, ppEntry);
}

HRESULT FilesInformationS3::BuildFromEntries(std::vector<Entry> entries) noexcept
{
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
