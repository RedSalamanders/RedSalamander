#include "WorkspaceDiscovery.h"

#include "RedConfigureBinaryFile.h"

#include <algorithm>
#include <cstdint>
#include <cwctype>
#include <string_view>

#include <objidl.h>
#include <XmlLite.h>

#pragma warning(push)
// WIL: C4625 (copy ctor deleted), C4626 (copy assign deleted), C5026, C5027
#pragma warning(disable : 4625 4626 5026 5027 28182)
#include <wil/com.h>
#include <wil/resource.h>
#pragma warning(pop)

#pragma comment(lib, "Ole32.lib")
#pragma comment(lib, "XmlLite.lib")

namespace
{
namespace fs = std::filesystem;

[[nodiscard]] wchar_t ToLowerAscii(wchar_t ch) noexcept
{
    return static_cast<wchar_t>(std::towlower(static_cast<wint_t>(ch)));
}

[[nodiscard]] bool EqualsIgnoreCase(std::wstring_view lhs, std::wstring_view rhs) noexcept
{
    if (lhs.size() != rhs.size())
    {
        return false;
    }

    for (size_t index = 0; index < lhs.size(); ++index)
    {
        if (ToLowerAscii(lhs[index]) != ToLowerAscii(rhs[index]))
        {
            return false;
        }
    }

    return true;
}

[[nodiscard]] bool EndsWithIgnoreCase(std::wstring_view text, std::wstring_view suffix) noexcept
{
    if (suffix.size() > text.size())
    {
        return false;
    }
    return EqualsIgnoreCase(text.substr(text.size() - suffix.size()), suffix);
}

[[nodiscard]] bool HasPathSegment(const fs::path& path, std::wstring_view segment) noexcept
{
    for (const auto& part : path)
    {
        if (EqualsIgnoreCase(part.native(), segment))
        {
            return true;
        }
    }
    return false;
}

[[nodiscard]] bool IsSpecsTestRunsPath(const fs::path& path) noexcept
{
    const fs::path parent = path.parent_path();
    return EqualsIgnoreCase(path.filename().native(), L"TestRuns") && ! parent.empty() && EqualsIgnoreCase(parent.filename().native(), L"Specs");
}

[[nodiscard]] bool ShouldSkipWorkspaceDirectory(const fs::path& path) noexcept
{
    const std::wstring name = path.filename().native();
    return EqualsIgnoreCase(name, L".build") || EqualsIgnoreCase(name, L".git") || EqualsIgnoreCase(name, L".vs") ||
           EqualsIgnoreCase(name, L"vcpkg_installed") || EqualsIgnoreCase(name, L"x64") || IsSpecsTestRunsPath(path);
}

[[nodiscard]] bool IsVcxprojPath(const fs::path& path) noexcept
{
    return EqualsIgnoreCase(path.extension().native(), L".vcxproj");
}

[[nodiscard]] bool IsRcPath(const fs::path& path) noexcept
{
    return EqualsIgnoreCase(path.extension().native(), L".rc");
}

[[nodiscard]] bool IsThemeJson5Path(const fs::path& path) noexcept
{
    return EndsWithIgnoreCase(path.filename().native(), L".theme.json5");
}

[[nodiscard]] fs::path ResolveProjectRelativePath(const fs::path& projectPath, std::wstring_view includePath)
{
    fs::path value(includePath);
    if (! value.is_absolute())
    {
        value = projectPath.parent_path() / value;
    }
    return value.lexically_normal();
}

[[nodiscard]] HRESULT CreateStreamForBytes(const std::vector<uint8_t>& bytes, wil::com_ptr<IStream>& outStream) noexcept
{
    outStream.reset();
    if (bytes.empty())
    {
        return HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
    }

    wil::unique_hglobal memory(::GlobalAlloc(GMEM_MOVEABLE, bytes.size()));
    if (! memory)
    {
        return E_OUTOFMEMORY;
    }

    void* lockedMemory = ::GlobalLock(memory.get());
    if (! lockedMemory)
    {
        return E_OUTOFMEMORY;
    }
    std::copy(bytes.begin(), bytes.end(), static_cast<uint8_t*>(lockedMemory));
    if (::GlobalUnlock(memory.get()) == FALSE)
    {
        const DWORD unlockError = ::GetLastError();
        if (unlockError != NO_ERROR)
        {
            return HRESULT_FROM_WIN32(unlockError);
        }
    }

    IStream* rawStream = nullptr;
    const HRESULT hr   = ::CreateStreamOnHGlobal(memory.get(), TRUE, &rawStream);
    if (FAILED(hr))
    {
        return hr;
    }

    static_cast<void>(memory.release());
    outStream.attach(rawStream);
    return S_OK;
}

[[nodiscard]] HRESULT ExtractResourceCompileIncludes(const fs::path& projectPath, std::vector<std::wstring>& outIncludes) noexcept
{
    outIncludes.clear();

    std::vector<uint8_t> bytes;
    if (const HRESULT hr = RedConfigure::ReadBinaryFile(projectPath, bytes); FAILED(hr))
    {
        return hr;
    }

    wil::com_ptr<IStream> stream;
    if (const HRESULT hr = CreateStreamForBytes(bytes, stream); FAILED(hr))
    {
        return hr;
    }

    wil::com_ptr<IXmlReader> reader;
    if (const HRESULT hr = ::CreateXmlReader(__uuidof(IXmlReader), reader.put_void(), nullptr); FAILED(hr))
    {
        return hr;
    }
    static_cast<void>(reader->SetProperty(XmlReaderProperty_DtdProcessing, DtdProcessing_Prohibit));
    if (const HRESULT hr = reader->SetInput(stream.get()); FAILED(hr))
    {
        return hr;
    }

    XmlNodeType nodeType = XmlNodeType_None;
    HRESULT readHr       = reader->Read(&nodeType);
    while (readHr == S_OK)
    {
        if (nodeType != XmlNodeType_Element)
        {
            readHr = reader->Read(&nodeType);
            continue;
        }

        const wchar_t* localName = nullptr;
        if (FAILED(reader->GetLocalName(&localName, nullptr)) || ! localName || ! EqualsIgnoreCase(localName, L"ResourceCompile"))
        {
            readHr = reader->Read(&nodeType);
            continue;
        }

        const HRESULT moveHr = reader->MoveToAttributeByName(L"Include", nullptr);
        if (moveHr == S_FALSE)
        {
            static_cast<void>(reader->MoveToElement());
            readHr = reader->Read(&nodeType);
            continue;
        }
        if (FAILED(moveHr))
        {
            return moveHr;
        }

        const wchar_t* includeValue = nullptr;
        if (SUCCEEDED(reader->GetValue(&includeValue, nullptr)) && includeValue && includeValue[0] != L'\0')
        {
            outIncludes.emplace_back(includeValue);
        }
        static_cast<void>(reader->MoveToElement());
        readHr = reader->Read(&nodeType);
    }

    return FAILED(readHr) ? readHr : S_OK;
}

void DiscoverSatelliteResourcePaths(const fs::path& projectDirectory, std::vector<fs::path>& outPaths) noexcept
{
    outPaths.clear();

    const fs::path langRoot = projectDirectory / L"Lang";
    std::error_code ec;
    if (! fs::exists(langRoot, ec) || ! fs::is_directory(langRoot, ec))
    {
        return;
    }

    fs::recursive_directory_iterator it(langRoot, fs::directory_options::skip_permission_denied, ec);
    const fs::recursive_directory_iterator end;
    for (; ! ec && it != end; it.increment(ec))
    {
        const fs::directory_entry& entry = *it;
        if (entry.is_directory(ec))
        {
            continue;
        }
        if (entry.is_regular_file(ec) && IsRcPath(entry.path()))
        {
            outPaths.emplace_back(entry.path().lexically_normal());
        }
    }

    std::sort(outPaths.begin(), outPaths.end());
}

void DiscoverThemeFiles(const fs::path& root, std::vector<RedConfigure::Workspace::ThemeFile>& outThemeFiles) noexcept
{
    outThemeFiles.clear();

    std::error_code ec;
    fs::recursive_directory_iterator it(root, fs::directory_options::skip_permission_denied, ec);
    const fs::recursive_directory_iterator end;
    for (; ! ec && it != end; it.increment(ec))
    {
        const fs::directory_entry& entry = *it;
        const fs::path path              = entry.path();
        if (entry.is_directory(ec))
        {
            if (ShouldSkipWorkspaceDirectory(path))
            {
                it.disable_recursion_pending();
            }
            continue;
        }

        if (entry.is_regular_file(ec) && IsThemeJson5Path(path))
        {
            outThemeFiles.push_back(RedConfigure::Workspace::ThemeFile{.path = path.lexically_normal()});
        }
    }

    std::sort(outThemeFiles.begin(), outThemeFiles.end(), [](const auto& lhs, const auto& rhs) noexcept { return lhs.path < rhs.path; });
}

void DiscoverResourceOwners(const fs::path& root, std::vector<RedConfigure::Workspace::ResourceOwner>& outOwners, std::vector<std::wstring>& outErrors) noexcept
{
    outOwners.clear();

    std::error_code ec;
    fs::recursive_directory_iterator it(root, fs::directory_options::skip_permission_denied, ec);
    const fs::recursive_directory_iterator end;
    for (; ! ec && it != end; it.increment(ec))
    {
        const fs::directory_entry& entry = *it;
        const fs::path path              = entry.path();
        if (entry.is_directory(ec))
        {
            if (ShouldSkipWorkspaceDirectory(path))
            {
                it.disable_recursion_pending();
            }
            continue;
        }

        if (! entry.is_regular_file(ec) || ! IsVcxprojPath(path))
        {
            continue;
        }

        std::vector<std::wstring> includes;
        if (const HRESULT hr = ExtractResourceCompileIncludes(path, includes); FAILED(hr))
        {
            outErrors.emplace_back(path.wstring());
            continue;
        }

        fs::path embeddedResourcePath;
        for (const std::wstring& includePath : includes)
        {
            const fs::path resolved = ResolveProjectRelativePath(path, includePath);
            if (IsRcPath(resolved) && ! HasPathSegment(resolved, L"Lang"))
            {
                embeddedResourcePath = resolved;
                break;
            }
        }

        if (embeddedResourcePath.empty())
        {
            continue;
        }

        RedConfigure::Workspace::ResourceOwner owner;
        owner.name                 = path.stem().wstring();
        owner.projectPath          = path.lexically_normal();
        owner.embeddedResourcePath = embeddedResourcePath;
        DiscoverSatelliteResourcePaths(path.parent_path(), owner.satelliteResourcePaths);
        outOwners.push_back(std::move(owner));
    }

    std::sort(outOwners.begin(), outOwners.end(), [](const auto& lhs, const auto& rhs) noexcept { return lhs.name < rhs.name; });
}
} // namespace

namespace RedConfigure::Workspace
{
HRESULT DiscoverWorkspace(const std::filesystem::path& root, WorkspaceScanResult& outResult) noexcept
{
    outResult      = {};
    outResult.root = root.lexically_normal();

    std::error_code ec;
    if (root.empty() || ! fs::exists(root, ec) || ! fs::is_directory(root, ec))
    {
        return HRESULT_FROM_WIN32(ERROR_PATH_NOT_FOUND);
    }

    DiscoverResourceOwners(root, outResult.resourceOwners, outResult.errors);
    DiscoverThemeFiles(root, outResult.themeFiles);
    return S_OK;
}
} // namespace RedConfigure::Workspace
