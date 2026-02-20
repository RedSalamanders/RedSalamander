#include "ChangeCase.h"

#include "framework.h"

#include <algorithm>
#include <cstdint>
#include <cwctype>
#include <filesystem>
#include <limits>
#include <optional>
#include <span>
#include <string>
#include <unordered_set>
#include <vector>

#include <wil/com.h>

#include "PlugInterfaces/FileSystem.h"

namespace
{
[[nodiscard]] std::wstring ToLower(std::wstring_view text) noexcept
{
    std::wstring out(text);
    if (! out.empty())
    {
        ::CharLowerBuffW(out.data(), static_cast<DWORD>(out.size()));
    }
    return out;
}

[[nodiscard]] std::wstring ToUpper(std::wstring_view text) noexcept
{
    std::wstring out(text);
    if (! out.empty())
    {
        ::CharUpperBuffW(out.data(), static_cast<DWORD>(out.size()));
    }
    return out;
}

[[nodiscard]] std::wstring ToMixed(std::wstring_view text) noexcept
{
    std::wstring out(text);
    if (out.empty())
    {
        return out;
    }

    ::CharLowerBuffW(out.data(), static_cast<DWORD>(out.size()));

    bool newWord = true;
    for (wchar_t& ch : out)
    {
        if (std::iswalnum(static_cast<wint_t>(ch)) == 0)
        {
            newWord = true;
            continue;
        }

        if (newWord)
        {
            ch      = static_cast<wchar_t>(std::towupper(static_cast<wint_t>(ch)));
            newWord = false;
        }
    }

    return out;
}

[[nodiscard]] bool ContainsPathSeparator(std::wstring_view text) noexcept
{
    return text.find(L'\\') != std::wstring_view::npos || text.find(L'/') != std::wstring_view::npos;
}

[[nodiscard]] size_t PathDepthKey(const std::filesystem::path& p) noexcept
{
    size_t depth = 0;
    const std::wstring& text = p.native();
    for (wchar_t ch : text)
    {
        if (ch == L'\\' || ch == L'/')
        {
            ++depth;
        }
    }
    return depth;
}

[[nodiscard]] bool IsDotOrDotDot(std::wstring_view name) noexcept
{
    return name == L"." || name == L"..";
}

[[nodiscard]] wchar_t GuessPreferredSeparator(std::wstring_view folder) noexcept
{
    const bool hasForward = folder.find(L'/') != std::wstring_view::npos;
    const bool hasBack    = folder.find(L'\\') != std::wstring_view::npos;
    if (hasForward && ! hasBack)
    {
        return L'/';
    }
    return L'\\';
}

[[nodiscard]] std::filesystem::path JoinFolderAndLeaf(const std::filesystem::path& folder, std::wstring_view leaf) noexcept
{
    if (folder.empty())
    {
        return std::filesystem::path(std::wstring(leaf));
    }

    const std::wstring_view folderText(folder.native());
    std::wstring result(folderText);

    const wchar_t sep = GuessPreferredSeparator(folderText);
    if (! result.empty())
    {
        const wchar_t last = result.back();
        if (last != L'\\' && last != L'/')
        {
            result.push_back(sep);
        }
    }

    if (! leaf.empty())
    {
        result.append(leaf);
    }

    return std::filesystem::path(std::move(result));
}

struct RenameOp final
{
    std::filesystem::path sourcePath;
    std::wstring newLeaf;
    size_t depth = 0;
};

[[nodiscard]] HRESULT RenameBatch(IFileSystem& fileSystem,
                                 std::span<const RenameOp> ops,
                                 FileSystemFlags flags,
                                 const FileSystemOptions* options,
                                 IFileSystemCallback* callback,
                                 void* cookie) noexcept
{
    if (ops.empty())
    {
        return S_OK;
    }

    uint64_t totalBytes64 = static_cast<uint64_t>(ops.size()) * static_cast<uint64_t>(sizeof(FileSystemRenamePair));
    for (const RenameOp& op : ops)
    {
        const std::wstring& sourceText = op.sourcePath.native();
        const size_t sourceLen         = sourceText.size();
        const size_t nameLen           = op.newLeaf.size();

        totalBytes64 += static_cast<uint64_t>((sourceLen + 1u) * sizeof(wchar_t));
        totalBytes64 += static_cast<uint64_t>((nameLen + 1u) * sizeof(wchar_t));
        if (totalBytes64 > static_cast<uint64_t>(std::numeric_limits<unsigned long>::max()))
        {
            return HRESULT_FROM_WIN32(ERROR_ARITHMETIC_OVERFLOW);
        }
    }

    FileSystemArenaOwner arenaOwner;
    const HRESULT initHr = arenaOwner.Initialize(static_cast<unsigned long>(totalBytes64));
    if (FAILED(initHr))
    {
        return initHr;
    }

    FileSystemArena* arena = arenaOwner.Get();
    auto* pairs            = static_cast<FileSystemRenamePair*>(AllocateFromFileSystemArena(
        arena, static_cast<unsigned long>(ops.size() * sizeof(FileSystemRenamePair)), static_cast<unsigned long>(alignof(FileSystemRenamePair))));
    if (! pairs)
    {
        return E_OUTOFMEMORY;
    }

    for (size_t i = 0; i < ops.size(); ++i)
    {
        const RenameOp& op            = ops[i];
        const std::wstring& source   = op.sourcePath.native();
        const std::wstring_view name = op.newLeaf;

        if (source.empty() || name.empty())
        {
            return E_INVALIDARG;
        }

        if (ContainsPathSeparator(name))
        {
            return HRESULT_FROM_WIN32(ERROR_INVALID_NAME);
        }

        const size_t sourceLen = source.size();
        const size_t nameLen   = name.size();

        if (sourceLen > static_cast<size_t>((std::numeric_limits<unsigned long>::max)()) - 1u ||
            nameLen > static_cast<size_t>((std::numeric_limits<unsigned long>::max)()) - 1u)
        {
            return HRESULT_FROM_WIN32(ERROR_ARITHMETIC_OVERFLOW);
        }

        auto* sourceBuf = static_cast<wchar_t*>(AllocateFromFileSystemArena(
            arena, static_cast<unsigned long>((sourceLen + 1u) * sizeof(wchar_t)), static_cast<unsigned long>(alignof(wchar_t))));
        if (! sourceBuf)
        {
            return E_OUTOFMEMORY;
        }

        auto* nameBuf = static_cast<wchar_t*>(AllocateFromFileSystemArena(
            arena, static_cast<unsigned long>((nameLen + 1u) * sizeof(wchar_t)), static_cast<unsigned long>(alignof(wchar_t))));
        if (! nameBuf)
        {
            return E_OUTOFMEMORY;
        }

        ::CopyMemory(sourceBuf, source.data(), sourceLen * sizeof(wchar_t));
        sourceBuf[sourceLen] = L'\0';

        ::CopyMemory(nameBuf, name.data(), nameLen * sizeof(wchar_t));
        nameBuf[nameLen] = L'\0';

        pairs[i].sourcePath = sourceBuf;
        pairs[i].newName    = nameBuf;
    }

    if (ops.size() > static_cast<size_t>((std::numeric_limits<unsigned long>::max)()))
    {
        return HRESULT_FROM_WIN32(ERROR_ARITHMETIC_OVERFLOW);
    }

    return fileSystem.RenameItems(pairs, static_cast<unsigned long>(ops.size()), flags, options, callback, cookie);
}
} // namespace

namespace ChangeCase
{
std::wstring TransformLeafName(std::wstring_view leafName, const Options& options) noexcept
{
    const std::filesystem::path leafPath{std::wstring(leafName)};
    const std::wstring stem = leafPath.stem().wstring();
    const std::wstring ext  = leafPath.extension().wstring(); // includes leading dot

    switch (options.target)
    {
        case ChangeTarget::WholeFilename:
        {
            switch (options.style)
            {
                case CaseStyle::Lower: return ToLower(leafName);
                case CaseStyle::Upper: return ToUpper(leafName);
                case CaseStyle::Mixed: return ToMixed(leafName);
                case CaseStyle::PartiallyMixed:
                {
                    std::wstring result = ToMixed(stem);
                    result.append(ToLower(ext));
                    return result;
                }
            }
            break;
        }
        case ChangeTarget::OnlyName:
        {
            std::wstring result;
            switch (options.style)
            {
                case CaseStyle::Lower: result = ToLower(stem); break;
                case CaseStyle::Upper: result = ToUpper(stem); break;
                case CaseStyle::Mixed: result = ToMixed(stem); break;
                case CaseStyle::PartiallyMixed: result = ToMixed(stem); break;
            }
            result.append(ext);
            return result;
        }
        case ChangeTarget::OnlyExtension:
        {
            if (ext.empty())
            {
                return std::wstring(leafName);
            }

            std::wstring newExt;
            switch (options.style)
            {
                case CaseStyle::Lower: newExt = ToLower(ext); break;
                case CaseStyle::Upper: newExt = ToUpper(ext); break;
                case CaseStyle::Mixed: newExt = ToMixed(ext); break;
                case CaseStyle::PartiallyMixed: newExt = ToLower(ext); break;
            }

            std::wstring result = stem;
            result.append(std::move(newExt));
            return result;
        }
    }

    return std::wstring(leafName);
}

HRESULT ApplyToPaths(IFileSystem& fileSystem,
                     const std::vector<std::filesystem::path>& inputPaths,
                     const Options& options,
                     std::stop_token stopToken,
                     ProgressCallback progress,
                     void* progressCookie) noexcept
{
    if (inputPaths.empty())
    {
        return S_OK;
    }

    ProgressUpdate progressUpdate{};

    std::vector<std::filesystem::path> paths;
    paths.reserve(inputPaths.size());

    std::unordered_set<std::wstring> seen;
    seen.reserve(inputPaths.size() * 2u);

    const auto addPath = [&](const std::filesystem::path& p)
    {
        if (p.empty())
        {
            return;
        }

        const std::wstring key = p.native();
        if (key.empty())
        {
            return;
        }

        if (seen.insert(key).second)
        {
            paths.push_back(p);
        }
    };

    for (const auto& p : inputPaths)
    {
        addPath(p);
    }

    if (options.includeSubdirs)
    {
        std::vector<std::filesystem::path> pending;
        pending.reserve(inputPaths.size());
        for (const auto& root : inputPaths)
        {
            pending.push_back(root);
        }

        while (! pending.empty())
        {
            if (stopToken.stop_requested())
            {
                return HRESULT_FROM_WIN32(ERROR_CANCELLED);
            }

            std::filesystem::path directory = std::move(pending.back());
            pending.pop_back();

            progressUpdate.phase       = ProgressUpdate::Phase::Enumerating;
            progressUpdate.currentPath = directory;
            if (progress)
            {
                progress(progressUpdate, progressCookie);
            }

            wil::com_ptr<IFilesInformation> info;
            const HRESULT readHr = fileSystem.ReadDirectoryInfo(directory.c_str(), info.addressof());
            if (FAILED(readHr) || ! info)
            {
                continue;
            }

            FileInfo* buffer = nullptr;
            HRESULT hr       = info->GetBuffer(&buffer);
            if (FAILED(hr))
            {
                return hr;
            }

            if (! buffer)
            {
                ++progressUpdate.scannedFolders;
                continue;
            }

            unsigned long bufferSize = 0;
            hr                       = info->GetBufferSize(&bufferSize);
            if (FAILED(hr))
            {
                return hr;
            }

            unsigned char* bytes = reinterpret_cast<unsigned char*>(buffer);
            unsigned long offset = 0;

            for (FileInfo* entry = buffer; entry;)
            {
                if (stopToken.stop_requested())
                {
                    return HRESULT_FROM_WIN32(ERROR_CANCELLED);
                }

                if ((entry->FileNameSize % sizeof(wchar_t)) != 0u)
                {
                    return E_INVALIDARG;
                }

                const size_t nameChars = static_cast<size_t>(entry->FileNameSize) / sizeof(wchar_t);
                std::wstring_view name(entry->FileName, nameChars);
                if (! name.empty() && ! IsDotOrDotDot(name))
                {
                    const std::filesystem::path child = JoinFolderAndLeaf(directory, name);
                    if (! child.empty())
                    {
                        const std::wstring key = child.native();
                        if (! key.empty() && seen.insert(key).second)
                        {
                            paths.push_back(child);

                            const DWORD attrs          = entry->FileAttributes;
                            const bool isDir           = (attrs & FILE_ATTRIBUTE_DIRECTORY) != 0;
                            const bool isReparsePoint  = (attrs & FILE_ATTRIBUTE_REPARSE_POINT) != 0;
                            if (isDir && ! isReparsePoint)
                            {
                                pending.push_back(child);
                            }
                        }
                    }
                }

                ++progressUpdate.scannedEntries;

                if (entry->NextEntryOffset == 0)
                {
                    break;
                }

                if (entry->NextEntryOffset > bufferSize - offset)
                {
                    return HRESULT_FROM_WIN32(ERROR_BAD_LENGTH);
                }

                offset += entry->NextEntryOffset;
                if (offset >= bufferSize)
                {
                    return HRESULT_FROM_WIN32(ERROR_BAD_LENGTH);
                }

                entry = reinterpret_cast<FileInfo*>(bytes + offset);
            }

            ++progressUpdate.scannedFolders;
        }
    }

    if (paths.size() > 1u)
    {
        std::ranges::sort(
            paths,
            [](const std::filesystem::path& a, const std::filesystem::path& b) noexcept
            {
                const size_t depthA = PathDepthKey(a);
                const size_t depthB = PathDepthKey(b);
                if (depthA != depthB)
                {
                    return depthA > depthB;
                }
                return a.native().size() > b.native().size();
            });
    }

    std::vector<RenameOp> renames;
    renames.reserve(paths.size());

    for (const auto& path : paths)
    {
        const std::filesystem::path leaf = path.filename();
        const std::wstring oldLeaf       = leaf.wstring();
        if (oldLeaf.empty())
        {
            continue;
        }

        const std::wstring newLeaf = TransformLeafName(oldLeaf, options);
        if (newLeaf == oldLeaf)
        {
            continue;
        }

        RenameOp op{};
        op.sourcePath = path;
        op.newLeaf    = std::move(newLeaf);
        op.depth      = PathDepthKey(path);
        renames.push_back(std::move(op));
    }

    progressUpdate.phase          = ProgressUpdate::Phase::Renaming;
    progressUpdate.currentPath.clear();
    progressUpdate.plannedRenames   = static_cast<uint64_t>(renames.size());
    progressUpdate.completedRenames = 0;
    if (progress)
    {
        progress(progressUpdate, progressCookie);
    }

    uint64_t completed = 0;
    size_t index       = 0;
    const FileSystemFlags flags = FILESYSTEM_FLAG_NONE;
    while (index < renames.size())
    {
        if (stopToken.stop_requested())
        {
            return HRESULT_FROM_WIN32(ERROR_CANCELLED);
        }

        const size_t depth = renames[index].depth;
        size_t end         = index + 1u;
        while (end < renames.size() && renames[end].depth == depth)
        {
            ++end;
        }

        constexpr size_t kBatchSize = 64u;
        for (size_t batchStart = index; batchStart < end; batchStart += kBatchSize)
        {
            if (stopToken.stop_requested())
            {
                return HRESULT_FROM_WIN32(ERROR_CANCELLED);
            }

            const size_t batchEnd = std::min(end, batchStart + kBatchSize);

            progressUpdate.currentPath = renames[batchStart].sourcePath;
            if (progress)
            {
                progress(progressUpdate, progressCookie);
            }

            const HRESULT hr = RenameBatch(
                fileSystem, std::span<const RenameOp>(renames.data() + batchStart, batchEnd - batchStart), flags, nullptr, nullptr, nullptr);
            if (FAILED(hr))
            {
                return hr;
            }

            completed += static_cast<uint64_t>(batchEnd - batchStart);
            progressUpdate.completedRenames = completed;
            progressUpdate.currentPath      = renames[batchEnd - 1u].sourcePath;
            if (progress)
            {
                progress(progressUpdate, progressCookie);
            }
        }

        index = end;
    }

    return S_OK;
}
} // namespace ChangeCase
