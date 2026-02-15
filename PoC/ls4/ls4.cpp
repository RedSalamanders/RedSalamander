#define WIN32_LEAN_AND_MEAN
#define NOMINMAX
#include <windows.h>

#include <algorithm>
#include <cstddef>
#include <cwchar>
#include <filesystem>
#include <format>
#include <memory>
#include <string>
#include <string_view>
#include <vector>

#pragma warning(push)
// WIL: C4625 (copy ctor deleted), C4626 (copy assign deleted), C5026 (move ctor deleted), C5027 (move assign deleted)
#pragma warning(disable : 4625 4626 5026 5027)
#include <wil/com.h>
#include <wil/resource.h>
#include <wil/win32_helpers.h>
#pragma warning(pop)

#include "PlugInterfaces/Factory.h"
#include "PlugInterfaces/FileSystem.h"

using CreateFactoryFunc = HRESULT(__stdcall*)(REFIID, const FactoryOptions*, IHost*, void**);

namespace
{
std::wstring FormatAttributes(unsigned long attributes)
{
    std::wstring buffer;
    auto appendFlag = [&buffer](std::wstring_view token)
    {
        if (! buffer.empty())
        {
            buffer.append(L"|");
        }
        buffer.append(token);
    };

    appendFlag((attributes & FILE_ATTRIBUTE_DIRECTORY) != 0U ? L"DIR" : L"FILE");
    if ((attributes & FILE_ATTRIBUTE_READONLY) != 0U)
    {
        appendFlag(L"READONLY");
    }
    if ((attributes & FILE_ATTRIBUTE_HIDDEN) != 0U)
    {
        appendFlag(L"HIDDEN");
    }
    if ((attributes & FILE_ATTRIBUTE_SYSTEM) != 0U)
    {
        appendFlag(L"SYSTEM");
    }
    if ((attributes & FILE_ATTRIBUTE_ARCHIVE) != 0U)
    {
        appendFlag(L"ARCHIVE");
    }
    if ((attributes & FILE_ATTRIBUTE_COMPRESSED) != 0U)
    {
        appendFlag(L"COMPRESSED");
    }
    if ((attributes & FILE_ATTRIBUTE_ENCRYPTED) != 0U)
    {
        appendFlag(L"ENCRYPTED");
    }
    if ((attributes & FILE_ATTRIBUTE_NOT_CONTENT_INDEXED) != 0U)
    {
        appendFlag(L"NOT_INDEXED");
    }
    if ((attributes & FILE_ATTRIBUTE_OFFLINE) != 0U)
    {
        appendFlag(L"OFFLINE");
    }
    if ((attributes & FILE_ATTRIBUTE_TEMPORARY) != 0U)
    {
        appendFlag(L"TEMP");
    }
    if ((attributes & FILE_ATTRIBUTE_REPARSE_POINT) != 0U)
    {
        appendFlag(L"REPARSE");
    }

    return buffer;
}

// format file size into most appropriate unit (B, KB, MB, GB)
std::wstring FormatSize(unsigned long long bytes)
{
    constexpr unsigned long long KB = 1024ULL;
    constexpr unsigned long long MB = KB * 1024ULL;
    constexpr unsigned long long GB = MB * 1024ULL;

    auto formatVal = [](double value, const wchar_t* unit) -> std::wstring
    {
        // for values < 10 in KB+ units keep one decimal, else no decimals
        if (value < 10.0 && unit[0] != L'B')
        {
            return std::format(L"{:.1f} {}", value, unit);
        }
        return std::format(L"{:.0f} {}", value, unit);
    };

    if (bytes >= GB)
    {
        return formatVal(static_cast<double>(bytes) / static_cast<double>(GB), L"GB");
    }
    if (bytes >= MB)
    {
        return formatVal(static_cast<double>(bytes) / static_cast<double>(MB), L"MB");
    }
    if (bytes >= KB)
    {
        return formatVal(static_cast<double>(bytes) / static_cast<double>(KB), L"KB");
    }
    return std::format(L"{} B", bytes);
}

void ReportError(std::wstring_view context, HRESULT hr)
{
    LPWSTR message     = nullptr;
    const DWORD flags  = FORMAT_MESSAGE_ALLOCATE_BUFFER | FORMAT_MESSAGE_FROM_SYSTEM | FORMAT_MESSAGE_IGNORE_INSERTS;
    const DWORD langId = MAKELANGID(LANG_NEUTRAL, SUBLANG_DEFAULT);

    const DWORD length = FormatMessageW(flags, nullptr, static_cast<DWORD>(hr), langId, reinterpret_cast<LPWSTR>(&message), 0, nullptr);
    wil::unique_hlocal localMessage(message);

    std::wstring description;
    if (length == 0 || message == nullptr)
    {
        description = std::format(L"0x{:08X}", static_cast<unsigned int>(hr));
    }
    else
    {
        description.assign(message, length);
        while (! description.empty() && (description.back() == L'\r' || description.back() == L'\n'))
        {
            description.pop_back();
        }
    }

    std::fwprintf(
        stderr, L"%.*ls failed: %ls (hr=0x%08X)\n", static_cast<int>(context.size()), context.data(), description.c_str(), static_cast<unsigned int>(hr));
}

HRESULT DisplayDirectory(IFilesInformation* filesInformation)
{
    if (filesInformation == nullptr)
    {
        return E_POINTER;
    }

    // Get buffer info
    unsigned long bufferSize    = 0;
    unsigned long allocatedSize = 0;
    HRESULT hr;
    hr = filesInformation->GetBufferSize(&bufferSize);
    if (FAILED(hr))
    {
        return hr;
    }
    hr = filesInformation->GetAllocatedSize(&allocatedSize);
    if (FAILED(hr))
    {
        return hr;
    }

    if (allocatedSize < bufferSize)
    {
        return HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
    }

    unsigned long count = 0;
    hr                  = filesInformation->GetCount(&count);
    if (FAILED(hr))
    {
        return hr;
    }

    FileInfo* buffer = nullptr;
    hr               = filesInformation->GetBuffer(&buffer);
    if (FAILED(hr))
    {
        return hr;
    }
    if (buffer == nullptr)
    {
        std::wprintf(L"\n========== BUFFER INFO ==========\n");
        std::wprintf(L"Buffer size: %ls\n", FormatSize(bufferSize).c_str());
        std::wprintf(L"Buffer allocated: %ls\n", FormatSize(allocatedSize).c_str());
        std::wprintf(L"Buffer utilization: %.1f%%\n", (static_cast<double>(bufferSize) / static_cast<double>(allocatedSize)) * 100.0);
        std::wprintf(L"Entry count: %lu\n", count);
        std::wprintf(L"=================================\n\n");
        return S_OK;
    }

    std::wprintf(L"\n========== BUFFER INFO ==========\n");
    std::wprintf(L"Buffer size: %ls\n", FormatSize(bufferSize).c_str());
    std::wprintf(L"Buffer allocated: %ls\n", FormatSize(allocatedSize).c_str());
    std::wprintf(L"Buffer utilization: %.1f%%\n", (static_cast<double>(bufferSize) / static_cast<double>(allocatedSize)) * 100.0);
    std::wprintf(L"Entry count: %lu\n", count);
    std::wprintf(L"=================================\n\n");

    unsigned long fileCount            = 0;
    unsigned long dirCount             = 0;
    unsigned long long totalSize       = 0;
    unsigned long long totalEntryBytes = 0;

    std::byte* current               = reinterpret_cast<std::byte*>(buffer);
    const std::byte* const bufferEnd = current + bufferSize;
    const size_t baseSize            = offsetof(FileInfo, FileName);

    for (unsigned long i = 0; i < count && current < bufferEnd; ++i)
    {
        if (static_cast<size_t>(bufferEnd - current) < baseSize)
        {
            continue;
        }

        auto* entry = reinterpret_cast<FileInfo*>(current);

        const size_t nameChars = entry->FileNameSize / sizeof(wchar_t);
        std::wstring name(entry->FileName, entry->FileName + nameChars);
        const std::wstring attributes = FormatAttributes(entry->FileAttributes);

        // Calculate file size
        const std::wstring sizeStr = FormatSize(static_cast<unsigned long long>(entry->EndOfFile));

        std::wprintf(L"%-40ls\t%12ls\t%ls\n", name.c_str(), sizeStr.c_str(), attributes.c_str());

        // Count files and directories
        if ((entry->FileAttributes & FILE_ATTRIBUTE_DIRECTORY) != 0U)
        {
            ++dirCount;
        }
        else
        {
            ++fileCount;
            totalSize += static_cast<unsigned long long>(entry->EndOfFile);
        }

        // Calculate entry size in buffer
        const size_t entrySize = baseSize + entry->FileNameSize + sizeof(wchar_t);
        // Align to 4 bytes
        const size_t alignedSize = (entrySize + 3) & ~3;
        totalEntryBytes += alignedSize;

        const size_t stride = entry->NextEntryOffset != 0 ? static_cast<size_t>(entry->NextEntryOffset) : alignedSize;
        if (stride == 0 || stride > static_cast<size_t>(bufferEnd - current))
        {
            break;
        }
        current += stride;
    }

    // Display summary
    std::wprintf(L"\n");
    std::wprintf(L"========== SUMMARY ==========\n");
    std::wprintf(L"Total items:        %lu (%lu file(s), %lu dir(s))\n", fileCount + dirCount, fileCount, dirCount);
    std::wprintf(L"Total file size:    %ls\n", FormatSize(totalSize).c_str());
    std::wprintf(L"Buffer allocated:        %ls\n", FormatSize(allocatedSize).c_str());
    std::wprintf(L"Buffer used:        %ls\n", FormatSize(totalEntryBytes).c_str());
    std::wprintf(L"Buffer utilization: %.1f%%\n", (static_cast<double>(totalEntryBytes) / static_cast<double>(totalSize)) * 100.0);
    if (count > 0)
    {
        std::wprintf(L"Avg entry size:     %lu bytes\n", static_cast<unsigned long>(totalEntryBytes / count));
    }
    std::wprintf(L"=============================\n");

    return S_OK;
}
} // namespace

int wmain(int argc, wchar_t** argv)
{
    const std::wstring path = (argc >= 2 && argv[1] != nullptr && argv[1][0] != L'\0') ? std::wstring(argv[1]) : std::wstring(L".");

    wil::unique_cotaskmem_string modulePath;
    const HRESULT pathHr = wil::GetModuleFileNameW<wil::unique_cotaskmem_string>(nullptr, modulePath);
    if (FAILED(pathHr) || ! modulePath)
    {
        ReportError(L"GetModuleFileNameW", FAILED(pathHr) ? pathHr : E_FAIL);
        return 1;
    }

    const std::filesystem::path exeDir     = std::filesystem::path(modulePath.get()).parent_path();
    const std::filesystem::path pluginPath = exeDir / L"Plugins" / L"FileSystem.dll";

    wil::unique_hmodule library(LoadLibraryW(pluginPath.c_str()));
    if (! library)
    {
        ReportError(L"LoadLibraryW(Plugins\\FileSystem.dll)", HRESULT_FROM_WIN32(GetLastError()));
        return 1;
    }

#pragma warning(push)
#pragma warning(disable : 4191) // reinterpret_cast from FARPROC to function pointer
    const auto createFactory = reinterpret_cast<CreateFactoryFunc>(GetProcAddress(library.get(), "RedSalamanderCreate"));
#pragma warning(pop)
    if (! createFactory)
    {
        ReportError(L"GetProcAddress(RedSalamanderCreate)", HRESULT_FROM_WIN32(GetLastError()));
        return 2;
    }

    FactoryOptions options{};
    options.debugLevel = DEBUG_LEVEL_NONE;

    wil::com_ptr_nothrow<IFileSystem> fileSystem;
    HRESULT hr = createFactory(__uuidof(IFileSystem), &options, nullptr, fileSystem.put_void());
    if (FAILED(hr))
    {
        ReportError(L"RedSalamanderCreate(FileSystem)", hr);
        return static_cast<int>(hr);
    }

    wil::com_ptr_nothrow<IFilesInformation> filesInformation;
    // Fill filesInformation with directory info
    hr = fileSystem->ReadDirectoryInfo(path.c_str(), filesInformation.put());
    if (FAILED(hr))
    {
        ReportError(L"ReadDirectoryInfo", hr);
        return static_cast<int>(hr);
    }

    hr = DisplayDirectory(filesInformation.get());
    if (FAILED(hr))
    {
        ReportError(L"DisplayDirectory", hr);
        return static_cast<int>(hr);
    }

    return 0;
}
