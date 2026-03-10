#pragma once

#include "PlugInterfaces/FileSystem.h"

namespace SearchFallbackEngine
{
HRESULT Execute(IFileSystem* fileSystem, const FileSystemSearchQuery* query, IFileSystemSearchCallback* callback, void* cookie) noexcept;
}
