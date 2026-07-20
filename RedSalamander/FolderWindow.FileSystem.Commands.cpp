#include "ChangeCase.h"
#include "ConnectionManagerWindow.h"
#include "ConnectionSecrets.h"
#include "DxUi/DxUi.h"
#include "DxUiThemePalette.h"
#include "FileActionLauncher.h"
#include "FileActionResolver.h"
#include "FolderWindow.FileSystem.Private.h"
#include "FolderWindowInternal.h"
#include "Helpers.h"
#include "HostServices.h"
#include "LocalFileTransaction.h"
#include "MaskSyntax.h"
#include "NavigationLocation.h"
#include "PathUtils.h"
#include "SettingsStore.h"
#include "ViewerPluginManager.h"
#include "Win32CallbackHelpers.h"
#include "WindowMessages.h"
#include "WindowSizing.h"
#ifdef ENABLE_TESTS
#include "SelfTestCommon.h"
#endif
#include "UiMetrics.h"
#include "UnicodeClipboard.h"

#include <algorithm>
#include <array>
#include <atomic>
#include <chrono>
#include <cstddef>
#include <cstdlib>
#include <cstring>
#include <cwctype>
#include <fstream>
#include <limits>
#include <map>
#include <span>
#include <string>
#include <system_error>
#include <unordered_set>
#include <utility>
#include <vector>

#include <commdlg.h>
#include <lm.h>
#include <oleauto.h>
#include <shellapi.h>
#include <shlwapi.h>
#include <shobjidl.h>
#include <winnetwk.h>

#ifndef INITGUID
#define INITGUID
#define REDSALAMANDER_UNDEF_7ZIP_INITGUID
#endif
#include <7zip/CPP/7zip/Archive/IArchive.h>
#include <7zip/CPP/7zip/IPassword.h>
#ifdef REDSALAMANDER_UNDEF_7ZIP_INITGUID
#undef INITGUID
#undef REDSALAMANDER_UNDEF_7ZIP_INITGUID
#endif

#pragma comment(lib, "Comdlg32.lib")
#pragma comment(lib, "Netapi32.lib")
#pragma comment(lib, "OleAut32.lib")
#pragma comment(lib, "Shlwapi.lib")

#pragma warning(push)
#pragma warning(disable : 6297 28182)
#include <yyjson.h>
#pragma warning(pop)

#pragma warning(push)
#pragma warning(disable : 4625 4626 5026 5027)
#include <wil/resource.h>
#pragma warning(pop)

using namespace FolderWindowFileSystemInternal;
using OrdinalString::EqualsNoCase;

namespace
{
[[nodiscard]] std::wstring BuildCreateDirectorySuggestedName(std::wstring_view baseName, unsigned int suffix)
{
    if (suffix == 0u)
    {
        return std::wstring(baseName);
    }

    return std::format(L"{} ({})", baseName, suffix);
}

[[nodiscard]] bool DirectoryNamesMatch(std::wstring_view left, std::wstring_view right, bool ignoreCase) noexcept
{
    if (! ignoreCase)
    {
        return left == right;
    }

    return CompareStringOrdinal(left.data(), static_cast<int>(left.size()), right.data(), static_cast<int>(right.size()), TRUE) == CSTR_EQUAL;
}

[[nodiscard]] bool TryCollectExistingDirectoryNames(const wil::com_ptr<IFileSystem>& fileSystem,
                                                    const std::filesystem::path& folder,
                                                    std::vector<std::wstring>& outNames) noexcept
{
    outNames.clear();
    if (! fileSystem)
    {
        return false;
    }

    wil::com_ptr<IFilesInformation> filesInformation;
    const HRESULT hr = fileSystem->ReadDirectoryInfo(folder.c_str(), filesInformation.put());
    if (FAILED(hr) || ! filesInformation)
    {
        return false;
    }

    FileInfo* entry = nullptr;
    if (FAILED(filesInformation->GetBuffer(&entry)) || entry == nullptr)
    {
        return true;
    }

    while (entry != nullptr)
    {
        if ((entry->FileAttributes & FILE_ATTRIBUTE_DIRECTORY) != 0)
        {
            const size_t nameChars = static_cast<size_t>(entry->FileNameSize) / sizeof(wchar_t);
            const std::wstring_view name(entry->FileName, nameChars);
            if (name != L"." && name != L"..")
            {
                outNames.emplace_back(name);
            }
        }

        if (entry->NextEntryOffset == 0)
        {
            break;
        }

        entry = reinterpret_cast<FileInfo*>(reinterpret_cast<std::byte*>(entry) + entry->NextEntryOffset);
    }

    return true;
}

[[nodiscard]] std::wstring ResolveInitialCreateDirectoryName(const wil::com_ptr<IFileSystem>& fileSystem,
                                                             const std::filesystem::path& folder,
                                                             std::wstring_view defaultName,
                                                             bool ignoreCase) noexcept
{
    std::vector<std::wstring> existingDirectoryNames;
    if (! TryCollectExistingDirectoryNames(fileSystem, folder, existingDirectoryNames))
    {
        return std::wstring(defaultName);
    }

    const auto isTaken = [&](std::wstring_view candidate) noexcept
    {
        for (const std::wstring& existingName : existingDirectoryNames)
        {
            if (DirectoryNamesMatch(existingName, candidate, ignoreCase))
            {
                return true;
            }
        }

        return false;
    };

    for (unsigned int suffix = 0u; suffix < 10000u; ++suffix)
    {
        std::wstring candidate = BuildCreateDirectorySuggestedName(defaultName, suffix);
        if (! isTaken(candidate))
        {
            return candidate;
        }
    }

    return std::format(L"{} ({})", defaultName, GetTickCount64());
}

[[nodiscard]] bool TryParseCreateDirectorySuggestedSuffix(std::wstring_view requestedName, std::wstring_view defaultName, unsigned int& outSuffix) noexcept
{
    if (requestedName == defaultName)
    {
        outSuffix = 0u;
        return true;
    }

    std::wstring prefix(defaultName);
    prefix.append(L" (");
    if (requestedName.size() <= prefix.size() || ! requestedName.starts_with(prefix) || requestedName.back() != L')')
    {
        return false;
    }

    const std::wstring_view digits = requestedName.substr(prefix.size(), requestedName.size() - prefix.size() - 1u);
    if (digits.empty())
    {
        return false;
    }

    uint64_t parsedValue = 0u;
    for (const wchar_t ch : digits)
    {
        if (ch < L'0' || ch > L'9')
        {
            return false;
        }

        parsedValue = (parsedValue * 10u) + static_cast<uint64_t>(ch - L'0');
        if (parsedValue > static_cast<uint64_t>((std::numeric_limits<unsigned int>::max)()))
        {
            return false;
        }
    }

    if (parsedValue == 0u)
    {
        return false;
    }

    outSuffix = static_cast<unsigned int>(parsedValue);
    return true;
}

using ShellNewTemplateDefinition = FolderWindow::ShellNewTemplateDefinition;
using ShellNewTemplateKind       = FolderWindow::ShellNewTemplateKind;

[[nodiscard]] std::wstring LowerShellNewTemplateId(std::wstring_view text)
{
    std::wstring id;
    id.reserve(text.size());
    for (const wchar_t ch : text)
    {
        if ((ch >= L'a' && ch <= L'z') || (ch >= L'0' && ch <= L'9') || ch == L'_' || ch == L'-')
        {
            id.push_back(ch);
        }
        else if (ch >= L'A' && ch <= L'Z')
        {
            id.push_back(static_cast<wchar_t>(std::towlower(static_cast<wint_t>(ch))));
        }
        else
        {
            id.push_back(L'_');
        }
    }
    return id;
}

[[nodiscard]] std::wstring TrimNullTerminatedRegistryString(std::wstring text)
{
    while (! text.empty() && text.back() == L'\0')
    {
        text.pop_back();
    }
    return StringUtils::TrimWhitespaceCopy(text);
}

[[nodiscard]] std::optional<std::wstring> ReadRegistryStringValue(HKEY key, const wchar_t* valueName)
{
    DWORD type               = 0;
    DWORD byteCount          = 0;
    const LSTATUS sizeStatus = RegQueryValueExW(key, valueName, nullptr, &type, nullptr, &byteCount);
    if (sizeStatus != ERROR_SUCCESS || (type != REG_SZ && type != REG_EXPAND_SZ) || byteCount == 0u)
    {
        return std::nullopt;
    }

    std::wstring value((byteCount / sizeof(wchar_t)) + 1u, L'\0');
    const LSTATUS readStatus = RegQueryValueExW(key, valueName, nullptr, &type, reinterpret_cast<BYTE*>(value.data()), &byteCount);
    if (readStatus != ERROR_SUCCESS || (type != REG_SZ && type != REG_EXPAND_SZ))
    {
        return std::nullopt;
    }

    value.resize(byteCount / sizeof(wchar_t));
    value = TrimNullTerminatedRegistryString(std::move(value));
    if (value.empty())
    {
        return std::nullopt;
    }

    if (type == REG_EXPAND_SZ)
    {
        const DWORD expandedChars = ExpandEnvironmentStringsW(value.c_str(), nullptr, 0);
        if (expandedChars > 0u)
        {
            std::wstring expanded(expandedChars, L'\0');
            const DWORD written = ExpandEnvironmentStringsW(value.c_str(), expanded.data(), expandedChars);
            if (written > 0u && written <= expandedChars)
            {
                value = TrimNullTerminatedRegistryString(std::move(expanded));
            }
        }
    }

    return value.empty() ? std::nullopt : std::optional<std::wstring>(std::move(value));
}

[[nodiscard]] std::optional<std::vector<std::byte>> ReadRegistryBinaryValue(HKEY key, const wchar_t* valueName)
{
    DWORD type               = 0;
    DWORD byteCount          = 0;
    const LSTATUS sizeStatus = RegQueryValueExW(key, valueName, nullptr, &type, nullptr, &byteCount);
    if (sizeStatus != ERROR_SUCCESS || type != REG_BINARY)
    {
        return std::nullopt;
    }

    std::vector<std::byte> value(byteCount);
    if (byteCount == 0u)
    {
        return value;
    }

    const LSTATUS readStatus = RegQueryValueExW(key, valueName, nullptr, &type, reinterpret_cast<BYTE*>(value.data()), &byteCount);
    if (readStatus != ERROR_SUCCESS || type != REG_BINARY)
    {
        return std::nullopt;
    }
    value.resize(byteCount);
    return value;
}

[[nodiscard]] bool RegistryValueExists(HKEY key, const wchar_t* valueName) noexcept
{
    DWORD type = 0;
    return RegQueryValueExW(key, valueName, nullptr, &type, nullptr, nullptr) == ERROR_SUCCESS;
}

[[nodiscard]] std::wstring LoadIndirectShellString(std::wstring_view raw)
{
    std::wstring text(raw);
    if (text.empty() || text.front() != L'@')
    {
        return text;
    }

    std::wstring resolved(1024u, L'\0');
    const HRESULT hr = SHLoadIndirectString(text.c_str(), resolved.data(), static_cast<UINT>(resolved.size()), nullptr);
    if (FAILED(hr))
    {
        return text;
    }
    return TrimNullTerminatedRegistryString(std::move(resolved));
}

[[nodiscard]] std::wstring ReadFileTypeDisplayName(std::wstring_view extension)
{
    std::wstring subKey(extension);
    wil::unique_hkey extensionKey;
    if (RegOpenKeyExW(HKEY_CLASSES_ROOT, subKey.c_str(), 0, KEY_READ, extensionKey.put()) != ERROR_SUCCESS)
    {
        return {};
    }

    std::optional<std::wstring> progId = ReadRegistryStringValue(extensionKey.get(), nullptr);
    if (! progId.has_value())
    {
        return {};
    }

    wil::unique_hkey progKey;
    if (RegOpenKeyExW(HKEY_CLASSES_ROOT, progId->c_str(), 0, KEY_READ, progKey.put()) != ERROR_SUCCESS)
    {
        return {};
    }

    if (std::optional<std::wstring> friendlyType = ReadRegistryStringValue(progKey.get(), L"FriendlyTypeName"); friendlyType.has_value())
    {
        std::wstring loaded = LoadIndirectShellString(friendlyType.value());
        if (! loaded.empty())
        {
            return loaded;
        }
    }

    if (std::optional<std::wstring> defaultName = ReadRegistryStringValue(progKey.get(), nullptr); defaultName.has_value())
    {
        return defaultName.value();
    }

    return {};
}

[[nodiscard]] std::wstring SanitizeShellNewFileNamePart(std::wstring_view raw)
{
    std::wstring text = StringUtils::TrimWhitespaceCopy(raw);
    if (text.empty())
    {
        return L"New File";
    }

    for (wchar_t& ch : text)
    {
        if (ch < 32 || ch == L'<' || ch == L'>' || ch == L':' || ch == L'"' || ch == L'/' || ch == L'\\' || ch == L'|' || ch == L'?' || ch == L'*')
        {
            ch = L'_';
        }
    }
    while (! text.empty() && (text.back() == L'.' || text.back() == L' '))
    {
        text.pop_back();
    }
    return text.empty() ? std::wstring(L"New File") : text;
}

[[nodiscard]] bool EndsWithNoCaseText(std::wstring_view text, std::wstring_view suffix) noexcept
{
    if (suffix.size() > text.size())
    {
        return false;
    }

    const std::wstring_view tail = text.substr(text.size() - suffix.size());
    return CompareStringOrdinal(tail.data(), static_cast<int>(tail.size()), suffix.data(), static_cast<int>(suffix.size()), TRUE) == CSTR_EQUAL;
}

[[nodiscard]] std::wstring BuildShellNewDefaultFileName(std::wstring_view displayName, std::wstring_view extension)
{
    std::wstring name = SanitizeShellNewFileNamePart(displayName);
    if (! name.starts_with(L"New "))
    {
        name = L"New " + name;
    }
    if (! extension.empty() && ! EndsWithNoCaseText(name, extension))
    {
        name.append(extension);
    }
    return name;
}

[[nodiscard]] std::optional<std::filesystem::path> ResolveShellNewFileNameTemplatePath(std::wstring_view rawFileName)
{
    std::wstring fileName(rawFileName);
    const DWORD expandedChars = ExpandEnvironmentStringsW(fileName.c_str(), nullptr, 0);
    if (expandedChars > 0u)
    {
        std::wstring expanded(expandedChars, L'\0');
        const DWORD written = ExpandEnvironmentStringsW(fileName.c_str(), expanded.data(), expandedChars);
        if (written > 0u && written <= expandedChars)
        {
            fileName = TrimNullTerminatedRegistryString(std::move(expanded));
        }
    }

    std::error_code ec;
    std::filesystem::path candidate(fileName);
    if (candidate.is_absolute() && std::filesystem::exists(candidate, ec))
    {
        return candidate;
    }
    ec.clear();

    std::wstring windowsDirectory(MAX_PATH, L'\0');
    const UINT chars = GetWindowsDirectoryW(windowsDirectory.data(), static_cast<UINT>(windowsDirectory.size()));
    if (chars == 0u || chars >= windowsDirectory.size())
    {
        return std::nullopt;
    }
    windowsDirectory.resize(chars);

    candidate = std::filesystem::path(windowsDirectory) / L"ShellNew" / fileName;
    if (std::filesystem::exists(candidate, ec))
    {
        return candidate;
    }
    return std::nullopt;
}

[[nodiscard]] std::optional<ShellNewTemplateDefinition> TryBuildShellNewTemplateFromRegistry(std::wstring_view extension, HKEY shellNewKey)
{
    std::wstring displayName;
    if (std::optional<std::wstring> itemName = ReadRegistryStringValue(shellNewKey, L"ItemName"); itemName.has_value())
    {
        displayName = LoadIndirectShellString(itemName.value());
    }
    if (displayName.empty())
    {
        displayName = ReadFileTypeDisplayName(extension);
    }
    if (displayName.empty())
    {
        displayName = extension.size() > 1u ? std::wstring(extension.substr(1)) + L" File" : std::wstring(L"File");
    }

    ShellNewTemplateDefinition result{};
    result.id              = LowerShellNewTemplateId(extension.size() > 1u && extension.front() == L'.' ? extension.substr(1) : extension);
    result.displayName     = displayName;
    result.extension       = std::wstring(extension);
    result.defaultFileName = BuildShellNewDefaultFileName(result.displayName, result.extension);

    if (std::optional<std::vector<std::byte>> data = ReadRegistryBinaryValue(shellNewKey, L"Data"); data.has_value())
    {
        result.kind = ShellNewTemplateKind::Data;
        result.data = std::move(data.value());
        return result;
    }

    if (std::optional<std::wstring> fileName = ReadRegistryStringValue(shellNewKey, L"FileName"); fileName.has_value())
    {
        if (std::optional<std::filesystem::path> templatePath = ResolveShellNewFileNameTemplatePath(fileName.value()); templatePath.has_value())
        {
            result.kind             = ShellNewTemplateKind::FileName;
            result.templateFilePath = templatePath.value();
            return result;
        }
    }

    if (RegistryValueExists(shellNewKey, L"NullFile"))
    {
        result.kind = ShellNewTemplateKind::NullFile;
        return result;
    }

    return std::nullopt;
}

[[nodiscard]] std::vector<ShellNewTemplateDefinition> EnumerateShellNewTemplatesFromRegistry()
{
    const auto startedAt = std::chrono::steady_clock::now();
    std::vector<ShellNewTemplateDefinition> templates;
    DWORD subKeyCount      = 0;
    DWORD maxSubKeyChars   = 0;
    uint64_t extensionKeys = 0u;

    if (RegQueryInfoKeyW(HKEY_CLASSES_ROOT, nullptr, nullptr, nullptr, &subKeyCount, &maxSubKeyChars, nullptr, nullptr, nullptr, nullptr, nullptr, nullptr) !=
        ERROR_SUCCESS)
    {
        Debug::Perf::Emit(L"shellnew.enumerate_us", L"registry", Debug::Perf::ElapsedUs(startedAt), 0u, 0u, E_FAIL);
        return templates;
    }

    std::wstring subKeyName(static_cast<size_t>(maxSubKeyChars) + 2u, L'\0');
    for (DWORD index = 0; index < subKeyCount; ++index)
    {
        DWORD charCount = static_cast<DWORD>(subKeyName.size());
        FILETIME writeTime{};
        const LSTATUS enumStatus = RegEnumKeyExW(HKEY_CLASSES_ROOT, index, subKeyName.data(), &charCount, nullptr, nullptr, nullptr, &writeTime);
        if (enumStatus != ERROR_SUCCESS)
        {
            continue;
        }

        std::wstring_view extension(subKeyName.data(), charCount);
        if (extension.empty() || extension.front() != L'.')
        {
            continue;
        }
        ++extensionKeys;

        std::wstring shellNewPath(extension);
        shellNewPath.append(L"\\ShellNew");
        wil::unique_hkey shellNewKey;
        if (RegOpenKeyExW(HKEY_CLASSES_ROOT, shellNewPath.c_str(), 0, KEY_READ, shellNewKey.put()) != ERROR_SUCCESS)
        {
            continue;
        }

        if (std::optional<ShellNewTemplateDefinition> entry = TryBuildShellNewTemplateFromRegistry(extension, shellNewKey.get()); entry.has_value())
        {
            templates.push_back(std::move(entry.value()));
        }
    }

    std::sort(templates.begin(),
              templates.end(),
              [](const ShellNewTemplateDefinition& left, const ShellNewTemplateDefinition& right) noexcept
    {
        const int nameCompare = CompareStringOrdinal(
            left.displayName.c_str(), static_cast<int>(left.displayName.size()), right.displayName.c_str(), static_cast<int>(right.displayName.size()), TRUE);
        if (nameCompare != CSTR_EQUAL)
        {
            return nameCompare == CSTR_LESS_THAN;
        }
        return CompareStringOrdinal(left.id.c_str(), static_cast<int>(left.id.size()), right.id.c_str(), static_cast<int>(right.id.size()), TRUE) ==
               CSTR_LESS_THAN;
    });

    Debug::Perf::Emit(L"shellnew.enumerate_us", L"registry", Debug::Perf::ElapsedUs(startedAt), extensionKeys, static_cast<uint64_t>(templates.size()), S_OK);
    return templates;
}

[[nodiscard]] std::wstring ShellNewTemplateKindDetail(ShellNewTemplateKind kind) noexcept
{
    switch (kind)
    {
        case ShellNewTemplateKind::NullFile: return L"null-file";
        case ShellNewTemplateKind::Data: return L"data";
        case ShellNewTemplateKind::FileName: return L"file-name";
    }
    return L"unknown";
}

[[nodiscard]] HRESULT WriteShellNewTemplateBytes(IFileWriter& writer, const void* data, size_t byteCount) noexcept
{
    const auto* cursor = static_cast<const std::byte*>(data);
    size_t remaining   = byteCount;
    while (remaining > 0u)
    {
        const unsigned long chunk = static_cast<unsigned long>((std::min)(remaining, static_cast<size_t>((std::numeric_limits<unsigned long>::max)())));
        unsigned long written     = 0;
        const HRESULT hr          = writer.Write(cursor, chunk, &written);
        if (FAILED(hr))
        {
            return hr;
        }
        if (written == 0u || written > chunk)
        {
            return HRESULT_FROM_WIN32(ERROR_WRITE_FAULT);
        }
        cursor += written;
        remaining -= written;
    }
    return S_OK;
}

[[nodiscard]] HRESULT WriteShellNewTemplateFile(IFileWriter& writer, const std::filesystem::path& sourcePath, uint64_t& bytesCopied) noexcept
{
    bytesCopied = 0u;
#pragma warning(push)
#pragma warning(disable : 4625 4626) // WIL unique_hfile copy operations are intentionally deleted.
    wil::unique_hfile sourceFile(
        CreateFileW(sourcePath.c_str(), GENERIC_READ, FILE_SHARE_READ | FILE_SHARE_DELETE, nullptr, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, nullptr));
#pragma warning(pop)
    if (! sourceFile)
    {
        return HRESULT_FROM_WIN32(GetLastError());
    }

    std::vector<std::byte> buffer(64u * 1024u);
    for (;;)
    {
        DWORD bytesRead = 0;
        if (ReadFile(sourceFile.get(), buffer.data(), static_cast<DWORD>(buffer.size()), &bytesRead, nullptr) == FALSE)
        {
            return HRESULT_FROM_WIN32(GetLastError());
        }
        if (bytesRead == 0u)
        {
            return S_OK;
        }
        const HRESULT writeHr = WriteShellNewTemplateBytes(writer, buffer.data(), bytesRead);
        if (FAILED(writeHr))
        {
            return writeHr;
        }
        bytesCopied += bytesRead;
    }
}

[[nodiscard]] std::wstring EnsureShellNewExtension(std::wstring requestedName, std::wstring_view extension)
{
    if (extension.empty())
    {
        return requestedName;
    }

    std::filesystem::path requestedPath(requestedName);
    if (requestedPath.extension().empty())
    {
        requestedName.append(extension);
    }
    return requestedName;
}

[[nodiscard]] std::wstring Utf16FromUtf8ForChangeAttributes(std::string_view text) noexcept
{
    return Common::Strings::Utf16FromUtf8StrictOrEmpty(text);
}

[[nodiscard]] bool IsSafeChangeAttributesStreamName(std::wstring_view streamName) noexcept
{
    if (streamName.empty())
    {
        return false;
    }

    for (const wchar_t ch : streamName)
    {
        if (ch == L':' || ch == L'\\' || ch == L'/' || ch == L'\0')
        {
            return false;
        }
    }

    return true;
}

[[nodiscard]] bool TryCollectRemovableStreamsFromPropertiesJson(std::string_view jsonUtf8, std::vector<std::wstring>& outStreams) noexcept
{
    outStreams.clear();
    if (jsonUtf8.empty())
    {
        return true;
    }

    std::string jsonCopy(jsonUtf8);
    yyjson_read_err err{};
    yyjson_doc* doc = yyjson_read_opts(jsonCopy.data(), jsonCopy.size(), YYJSON_READ_JSON5 | YYJSON_READ_ALLOW_BOM, nullptr, &err);
    if (! doc)
    {
        return false;
    }
    const auto freeDoc = wil::scope_exit([&] { yyjson_doc_free(doc); });

    yyjson_val* root = yyjson_doc_get_root(doc);
    if (! root || ! yyjson_is_obj(root))
    {
        return false;
    }

    yyjson_val* streams = yyjson_obj_get(root, "streams");
    if (! streams || ! yyjson_is_arr(streams))
    {
        return true;
    }

    std::unordered_set<std::wstring> seen;
    const size_t streamCount = yyjson_arr_size(streams);
    outStreams.reserve(streamCount);
    for (size_t index = 0; index < streamCount; ++index)
    {
        yyjson_val* stream = yyjson_arr_get(streams, index);
        if (! stream || ! yyjson_is_obj(stream))
        {
            continue;
        }

        yyjson_val* canRemoveVal = yyjson_obj_get(stream, "canRemove");
        if (! canRemoveVal || ! yyjson_is_bool(canRemoveVal) || yyjson_get_bool(canRemoveVal) == 0)
        {
            continue;
        }

        yyjson_val* nameVal = yyjson_obj_get(stream, "name");
        if (! nameVal || ! yyjson_is_str(nameVal))
        {
            continue;
        }

        const char* nameUtf8 = yyjson_get_str(nameVal);
        if (! nameUtf8 || nameUtf8[0] == '\0')
        {
            continue;
        }

        std::wstring streamName = Utf16FromUtf8ForChangeAttributes(nameUtf8);
        if (! IsSafeChangeAttributesStreamName(streamName) || seen.contains(streamName))
        {
            continue;
        }

        seen.insert(streamName);
        outStreams.emplace_back(std::move(streamName));
    }

    return true;
}

[[nodiscard]] HRESULT CollectRemovableStreamsForChangeAttributes(IFileSystemIO* fileIo,
                                                                 const std::filesystem::path& path,
                                                                 std::vector<std::wstring>& outStreams) noexcept
{
    outStreams.clear();
    if (! fileIo)
    {
        return E_POINTER;
    }

    const auto startedAt = std::chrono::steady_clock::now();

    const char* jsonUtf8 = nullptr;
    const HRESULT hr     = fileIo->GetItemProperties(path.c_str(), &jsonUtf8);
    if (FAILED(hr))
    {
        Debug::Perf::Emit(L"fileattrs.stream_enumerate_us", L"", Debug::Perf::ElapsedUs(startedAt), 0u, 0u, hr);
        return hr;
    }

    if (! TryCollectRemovableStreamsFromPropertiesJson(jsonUtf8 ? std::string_view(jsonUtf8) : std::string_view{}, outStreams))
    {
        constexpr HRESULT invalidDataHr = HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
        Debug::Perf::Emit(L"fileattrs.stream_enumerate_us", L"", Debug::Perf::ElapsedUs(startedAt), 0u, 0u, invalidDataHr);
        return invalidDataHr;
    }

    Debug::Perf::Emit(L"fileattrs.stream_enumerate_us", L"", Debug::Perf::ElapsedUs(startedAt), static_cast<uint64_t>(outStreams.size()), 0u, S_OK);
    return S_OK;
}

void RecordChangeAttributesFailure(FolderWindow::ChangeAttributesReport& report, HRESULT hr) noexcept
{
    ++report.failures;
    if (SUCCEEDED(report.firstFailure))
    {
        report.firstFailure = FAILED(hr) ? hr : E_FAIL;
    }
}

constexpr wchar_t kChangeAttributesOptionsPromptClassName[] = L"RedSalamander.ChangeAttributesOptionsPrompt";
constexpr wchar_t kOpenedFilesWindowClassName[]             = L"RedSalamander.OpenedFilesWindow";
constexpr wchar_t kSharedDirectoriesWindowClassName[]       = L"RedSalamander.SharedDirectoriesWindow";

#ifdef ENABLE_TESTS
std::atomic<HWND> g_changeAttributesOptionsPromptWindow{nullptr};

enum class ChangeAttributesOptionsPromptDebugCommand : uintptr_t
{
    GetSnapshot = 1u,
    SetState,
    CycleArchive,
    Confirm,
    Cancel,
};

struct ChangeAttributesOptionsPromptDebugStatePayload final
{
    uint8_t readOnly                = 0u;
    uint8_t hidden                  = 0u;
    uint8_t system                  = 0u;
    uint8_t archive                 = 0u;
    bool removeAlternateDataStreams = false;
};

[[nodiscard]] UINT GetChangeAttributesOptionsPromptDebugMessage() noexcept
{
    static const UINT message = RegisterWindowMessageW(L"RedSalamander.ChangeAttributesOptionsPrompt.Debug");
    return message;
}
#endif

#ifdef ENABLE_TESTS
[[nodiscard]] FolderWindow::AttributeChangeState AttributeChangeStateFromRaw(uint8_t value) noexcept
{
    switch (value)
    {
        case static_cast<uint8_t>(FolderWindow::AttributeChangeState::Set): return FolderWindow::AttributeChangeState::Set;
        case static_cast<uint8_t>(FolderWindow::AttributeChangeState::Clear): return FolderWindow::AttributeChangeState::Clear;
        default: return FolderWindow::AttributeChangeState::LeaveUnchanged;
    }
}
#endif

void ApplyAttributeCheckboxState(RedSalamander::DxUi::Checkbox* checkbox, FolderWindow::AttributeChangeState state) noexcept
{
    if (! checkbox)
    {
        return;
    }

    checkbox->SetChecked(state == FolderWindow::AttributeChangeState::Set);
    checkbox->SetIndeterminate(state == FolderWindow::AttributeChangeState::LeaveUnchanged);
}

[[nodiscard]] FolderWindow::AttributeChangeState NextAttributeChangeState(FolderWindow::AttributeChangeState state) noexcept
{
    switch (state)
    {
        case FolderWindow::AttributeChangeState::LeaveUnchanged: return FolderWindow::AttributeChangeState::Set;
        case FolderWindow::AttributeChangeState::Set: return FolderWindow::AttributeChangeState::Clear;
        case FolderWindow::AttributeChangeState::Clear: return FolderWindow::AttributeChangeState::LeaveUnchanged;
    }

    return FolderWindow::AttributeChangeState::LeaveUnchanged;
}

[[nodiscard]] FILETIME FileTimeFromTicks(int64_t ticks) noexcept
{
    ULARGE_INTEGER value{};
    value.QuadPart = static_cast<ULONGLONG>(ticks);

    FILETIME fileTime{};
    fileTime.dwLowDateTime  = value.LowPart;
    fileTime.dwHighDateTime = value.HighPart;
    return fileTime;
}

[[nodiscard]] int64_t FileTimeTicksFromFileTime(const FILETIME& fileTime) noexcept
{
    ULARGE_INTEGER value{};
    value.LowPart  = fileTime.dwLowDateTime;
    value.HighPart = fileTime.dwHighDateTime;
    return static_cast<int64_t>(value.QuadPart);
}

[[nodiscard]] int64_t GetCurrentFileTimeTicksForChangeAttributes() noexcept
{
    FILETIME fileTime{};
    GetSystemTimeAsFileTime(&fileTime);
    return FileTimeTicksFromFileTime(fileTime);
}

[[nodiscard]] bool TryGetLocalSystemTimeFromFileTimeTicks(int64_t ticks, SYSTEMTIME& out) noexcept
{
    if (ticks <= 0)
    {
        return false;
    }

    const FILETIME utcFileTime = FileTimeFromTicks(ticks);
    FILETIME localFileTime{};
    if (FileTimeToLocalFileTime(&utcFileTime, &localFileTime) == FALSE)
    {
        return false;
    }

    return FileTimeToSystemTime(&localFileTime, &out) != FALSE;
}

[[nodiscard]] std::wstring FormatDateForChangeAttributes(int64_t ticks)
{
    SYSTEMTIME local{};
    if (! TryGetLocalSystemTimeFromFileTimeTicks(ticks, local))
    {
        static_cast<void>(TryGetLocalSystemTimeFromFileTimeTicks(GetCurrentFileTimeTicksForChangeAttributes(), local));
    }

    return std::format(L"{:02}/{:02}/{:04}", local.wDay, local.wMonth, local.wYear);
}

[[nodiscard]] std::wstring FormatTimeForChangeAttributes(int64_t ticks)
{
    SYSTEMTIME local{};
    if (! TryGetLocalSystemTimeFromFileTimeTicks(ticks, local))
    {
        static_cast<void>(TryGetLocalSystemTimeFromFileTimeTicks(GetCurrentFileTimeTicksForChangeAttributes(), local));
    }

    return std::format(L"{:02}:{:02}:{:02}", local.wHour, local.wMinute, local.wSecond);
}

[[nodiscard]] std::wstring TrimChangeAttributesText(std::wstring_view text)
{
    size_t first = 0;
    while (first < text.size() && iswspace(text[first]) != 0)
    {
        ++first;
    }

    size_t last = text.size();
    while (last > first && iswspace(text[last - 1u]) != 0)
    {
        --last;
    }

    return std::wstring(text.substr(first, last - first));
}

[[nodiscard]] bool TryParseFixedDigits(std::wstring_view text, size_t offset, size_t count, WORD& out) noexcept
{
    if (offset + count > text.size() || count == 0u)
    {
        return false;
    }

    unsigned value = 0;
    for (size_t index = 0; index < count; ++index)
    {
        const wchar_t ch = text[offset + index];
        if (ch < L'0' || ch > L'9')
        {
            return false;
        }

        value = (value * 10u) + static_cast<unsigned>(ch - L'0');
    }

    if (value > std::numeric_limits<WORD>::max())
    {
        return false;
    }

    out = static_cast<WORD>(value);
    return true;
}

[[nodiscard]] bool TryParseChangeAttributesDate(std::wstring_view rawText, SYSTEMTIME& local) noexcept
{
    const std::wstring text = TrimChangeAttributesText(rawText);
    if (text.size() != 10u)
    {
        return false;
    }

    WORD day   = 0;
    WORD month = 0;
    WORD year  = 0;
    if (text[2] == L'/' && text[5] == L'/')
    {
        if (! TryParseFixedDigits(text, 0, 2, day) || ! TryParseFixedDigits(text, 3, 2, month) || ! TryParseFixedDigits(text, 6, 4, year))
        {
            return false;
        }
    }
    else if (text[4] == L'-' && text[7] == L'-')
    {
        if (! TryParseFixedDigits(text, 0, 4, year) || ! TryParseFixedDigits(text, 5, 2, month) || ! TryParseFixedDigits(text, 8, 2, day))
        {
            return false;
        }
    }
    else
    {
        return false;
    }

    if (year < 1601 || month < 1 || month > 12 || day < 1 || day > 31)
    {
        return false;
    }

    local.wYear  = year;
    local.wMonth = month;
    local.wDay   = day;
    return true;
}

[[nodiscard]] bool TryParseChangeAttributesTime(std::wstring_view rawText, SYSTEMTIME& local) noexcept
{
    const std::wstring text = TrimChangeAttributesText(rawText);
    if (text.size() != 5u && text.size() != 8u)
    {
        return false;
    }

    WORD hour   = 0;
    WORD minute = 0;
    WORD second = 0;
    if (text[2] != L':' || ! TryParseFixedDigits(text, 0, 2, hour) || ! TryParseFixedDigits(text, 3, 2, minute))
    {
        return false;
    }
    if (text.size() == 8u && (text[5] != L':' || ! TryParseFixedDigits(text, 6, 2, second)))
    {
        return false;
    }

    if (hour > 23 || minute > 59 || second > 59)
    {
        return false;
    }

    local.wHour         = hour;
    local.wMinute       = minute;
    local.wSecond       = second;
    local.wMilliseconds = 0;
    return true;
}

[[nodiscard]] bool TryParseChangeAttributesDateTime(std::wstring_view dateText, std::wstring_view timeText, int64_t& outTicks) noexcept
{
    SYSTEMTIME local{};
    if (! TryParseChangeAttributesDate(dateText, local) || ! TryParseChangeAttributesTime(timeText, local))
    {
        return false;
    }

    SYSTEMTIME utc{};
    if (TzSpecificLocalTimeToSystemTime(nullptr, &local, &utc) == FALSE)
    {
        utc = local;
    }

    FILETIME utcFileTime{};
    if (SystemTimeToFileTime(&utc, &utcFileTime) == FALSE)
    {
        return false;
    }

    outTicks = FileTimeTicksFromFileTime(utcFileTime);
    return true;
}

class ChangeAttributesOptionsPromptWindow final
{
public:
    ChangeAttributesOptionsPromptWindow(const ChangeAttributesOptionsPromptWindow&)            = delete;
    ChangeAttributesOptionsPromptWindow& operator=(const ChangeAttributesOptionsPromptWindow&) = delete;
    ChangeAttributesOptionsPromptWindow(ChangeAttributesOptionsPromptWindow&&)                 = delete;
    ChangeAttributesOptionsPromptWindow& operator=(ChangeAttributesOptionsPromptWindow&&)      = delete;

    ChangeAttributesOptionsPromptWindow(HWND ownerWindow,
                                        const AppTheme& theme,
                                        FolderWindow::ChangeAttributesOptions initialOptions = {},
                                        bool allowSubdirectories                             = false) noexcept
        : _ownerWindow(GetOwnerWindowOrSelf(ownerWindow)),
          _restoreFocusWindow(ownerWindow && IsWindow(ownerWindow) != FALSE ? ownerWindow : nullptr),
          _theme(theme),
          _options(initialOptions),
          _allowSubdirectories(allowSubdirectories)
    {
        if (_ownerWindow && IsWindow(_ownerWindow) != FALSE)
        {
            const HWND focused = GetFocus();
            if (focused && IsWindow(focused) != FALSE && (focused == _ownerWindow || IsChild(_ownerWindow, focused) != FALSE))
            {
                _restoreFocusWindow = focused;
            }
            else if (! _restoreFocusWindow || IsWindow(_restoreFocusWindow) == FALSE ||
                     (_restoreFocusWindow != _ownerWindow && IsChild(_ownerWindow, _restoreFocusWindow) == FALSE))
            {
                _restoreFocusWindow = _ownerWindow;
            }
        }
        else if (! _restoreFocusWindow || IsWindow(_restoreFocusWindow) == FALSE)
        {
            _restoreFocusWindow = nullptr;
        }
    }

    [[nodiscard]] std::optional<FolderWindow::ChangeAttributesOptions> ShowModal() noexcept
    {
        const HRESULT classHr = EnsureWindowClass();
        if (FAILED(classHr))
        {
            return std::nullopt;
        }

        const DWORD style   = WS_CAPTION | WS_SYSMENU | WS_POPUP | WS_CLIPCHILDREN | WS_CLIPSIBLINGS;
        const DWORD exStyle = WS_EX_DLGMODALFRAME;
        const UINT dpi      = _ownerWindow && IsWindow(_ownerWindow) != FALSE ? GetDpiForWindow(_ownerWindow) : GetDpiForSystem();

        RECT bounds{0, 0, ScalePanePromptForDpi(dpi, 520), ScalePanePromptForDpi(dpi, 480)};
        if (AdjustWindowRectExForDpi(&bounds, style, FALSE, exStyle, dpi) == FALSE)
        {
            return std::nullopt;
        }

        const bool restoreOwnerEnabled = _ownerWindow && IsWindow(_ownerWindow) != FALSE && IsWindowEnabled(_ownerWindow) != FALSE;
        if (restoreOwnerEnabled)
        {
            EnableWindow(_ownerWindow, FALSE);
        }
        const auto restoreOwner = wil::scope_exit([this, restoreOwnerEnabled]() noexcept
        {
            if (restoreOwnerEnabled && _ownerWindow && IsWindow(_ownerWindow) != FALSE)
            {
                EnableWindow(_ownerWindow, TRUE);
                SetActiveWindow(_ownerWindow);
                const HWND restoreFocus = (_restoreFocusWindow && IsWindow(_restoreFocusWindow) != FALSE &&
                                           (_restoreFocusWindow == _ownerWindow || IsChild(_ownerWindow, _restoreFocusWindow) != FALSE))
                                              ? _restoreFocusWindow
                                              : _ownerWindow;
                SetFocus(restoreFocus);
            }
        });

        const std::wstring caption = LoadStringResource(nullptr, IDS_CMD_CHANGE_ATTRIBUTES);
        const HWND hwnd            = CreateWindowExW(exStyle,
                                                     kChangeAttributesOptionsPromptClassName,
                                                     caption.c_str(),
                                                     style,
                                                     CW_USEDEFAULT,
                                                     CW_USEDEFAULT,
                                                     bounds.right - bounds.left,
                                                     bounds.bottom - bounds.top,
                                                     _ownerWindow,
                                                     nullptr,
                                                     GetModuleHandleW(nullptr),
                                                     this);
        if (! hwnd)
        {
            return std::nullopt;
        }
        if (! _hWnd)
        {
            _hWnd.reset(hwnd);
        }

        static_cast<void>(Common::WindowSizing::CenterExistingWindowOnOwner(_hWnd.get(), _ownerWindow));
        static_cast<void>(_dxHost.PrimeForShow());
        ShowWindow(_hWnd.get(), SW_SHOWNORMAL);
        UpdateWindow(_hWnd.get());
        SetForegroundWindow(_hWnd.get());

        MSG msg{};
        while (! _done)
        {
            const BOOL getMessageResult = GetMessageW(&msg, nullptr, 0, 0);
            if (getMessageResult == -1)
            {
                return std::nullopt;
            }
            if (getMessageResult == 0)
            {
                _done = true;
                break;
            }

            TranslateMessage(&msg);
            DispatchMessageW(&msg);
        }

        return _result;
    }

    static LRESULT CALLBACK WndProc(HWND hwnd, UINT message, WPARAM wParam, LPARAM lParam) noexcept
    {
        if (message == WM_NCCREATE)
        {
            auto* cs   = reinterpret_cast<CREATESTRUCTW*>(lParam);
            auto* self = static_cast<ChangeAttributesOptionsPromptWindow*>(cs ? cs->lpCreateParams : nullptr);
            if (! self)
            {
                return FALSE;
            }

            SetWindowLongPtrW(hwnd, GWLP_USERDATA, reinterpret_cast<LONG_PTR>(self));
            if (! self->_hWnd)
            {
                self->_hWnd.reset(hwnd);
            }
            return TRUE;
        }

        auto* self = reinterpret_cast<ChangeAttributesOptionsPromptWindow*>(GetWindowLongPtrW(hwnd, GWLP_USERDATA));
        if (! self)
        {
            return DefWindowProcW(hwnd, message, wParam, lParam);
        }

        bool handled     = false;
        LRESULT dxResult = 0;
        if (message != WM_CREATE)
        {
            dxResult = self->_dxHost.HandleMessage(hwnd, message, wParam, lParam, handled);
        }
        if (handled)
        {
            if (message == WM_SIZE || message == WM_DPICHANGED)
            {
                self->Layout();
            }
            if (message == WM_NCDESTROY)
            {
#ifdef ENABLE_TESTS
                g_changeAttributesOptionsPromptWindow.store(nullptr);
#endif
                SetWindowLongPtrW(hwnd, GWLP_USERDATA, 0);
                self->_dxHost.Detach();
                if (self->_hWnd.get() == hwnd)
                {
                    static_cast<void>(self->_hWnd.release());
                }
                self->_done = true;
            }
            return dxResult;
        }

#ifdef ENABLE_TESTS
        if (message == GetChangeAttributesOptionsPromptDebugMessage())
        {
            return self->OnDebugCommand(static_cast<ChangeAttributesOptionsPromptDebugCommand>(wParam), lParam);
        }
#endif

        switch (message)
        {
            case WM_CREATE: return self->OnCreate(hwnd) ? 0 : -1;
            case WM_SIZE: self->Layout(); return 0;
            case WM_DPICHANGED:
            {
                const auto* suggested = reinterpret_cast<const RECT*>(lParam);
                if (suggested)
                {
                    SetWindowPos(hwnd,
                                 nullptr,
                                 suggested->left,
                                 suggested->top,
                                 suggested->right - suggested->left,
                                 suggested->bottom - suggested->top,
                                 SWP_NOZORDER | SWP_NOACTIVATE);
                }
                self->Layout();
                return 0;
            }
            case WM_ERASEBKGND: return 1;
            case WM_NCACTIVATE: ApplyTitleBarTheme(hwnd, self->_theme, wParam != FALSE); return DefWindowProcW(hwnd, message, wParam, lParam);
            case WM_CLOSE: self->Cancel(); return 0;
            case WM_NCDESTROY:
#ifdef ENABLE_TESTS
                g_changeAttributesOptionsPromptWindow.store(nullptr);
#endif
                self->_dxHost.Detach();
                if (self->_hWnd.get() == hwnd)
                {
                    self->_hWnd.release();
                }
                SetWindowLongPtrW(hwnd, GWLP_USERDATA, 0);
                self->_done = true;
                break;
        }

        return DefWindowProcW(hwnd, message, wParam, lParam);
    }

private:
    struct DateTimeRowControls final
    {
        RedSalamander::DxUi::Label* label       = nullptr;
        RedSalamander::DxUi::Checkbox* checkbox = nullptr;
        RedSalamander::DxUi::TextField* date    = nullptr;
        RedSalamander::DxUi::TextField* time    = nullptr;
    };

    [[nodiscard]] static HRESULT EnsureWindowClass() noexcept
    {
        static ATOM atom = 0;
        if (atom != 0)
        {
            return S_OK;
        }

        WNDCLASSEXW wc{};
        wc.cbSize        = sizeof(wc);
        wc.lpfnWndProc   = ChangeAttributesOptionsPromptWindow::WndProc;
        wc.hInstance     = GetModuleHandleW(nullptr);
        wc.hCursor       = LoadCursorW(nullptr, IDC_ARROW);
        wc.lpszClassName = kChangeAttributesOptionsPromptClassName;
        wc.style         = CS_DBLCLKS;

        atom = RegisterClassExW(&wc);
        return atom != 0 ? S_OK : HRESULT_FROM_WIN32(GetLastError());
    }

    [[nodiscard]] bool OnCreate(HWND hwnd) noexcept
    {
        if (! _dxHost.Attach(hwnd))
        {
            return false;
        }

#ifdef ENABLE_TESTS
        g_changeAttributesOptionsPromptWindow.store(hwnd);
#endif
        BuildUi();
        ApplyTheme();
        Layout();
        if (_archiveCheckbox)
        {
            _dxHost.SetFocusControl(_archiveCheckbox);
        }
        _dxHost.SetDefaultButton(_okButton);
        _dxHost.SetCancelButton(_cancelButton);
        return true;
    }

    void BuildUi()
    {
        if (_root)
        {
            return;
        }

        using namespace RedSalamander::DxUi;

        _rootStorage = std::make_unique<Panel>();
        _root        = _rootStorage.get();

        _attributesSectionLabel = _root->AddChild<Label>(LoadStringResource(nullptr, IDS_CHANGE_ATTR_SECTION_ATTRIBUTES));
        _attributesSectionLabel->SetAlignment(DWRITE_TEXT_ALIGNMENT_LEADING);
        _attributesSectionLabel->SetFontRole(FontRole::BodyStrong);

        _archiveCheckbox    = BuildAttributeCheckbox(LoadStringResource(nullptr, IDS_CHANGE_ATTR_LABEL_ARCHIVE), _options.archive);
        _hiddenCheckbox     = BuildAttributeCheckbox(LoadStringResource(nullptr, IDS_CHANGE_ATTR_LABEL_HIDDEN), _options.hidden);
        _compressedCheckbox = _root->AddChild<Checkbox>(LoadStringResource(nullptr, IDS_CHANGE_ATTR_LABEL_COMPRESSED));
        _compressedCheckbox->SetEnabled(false);

        _readOnlyCheckbox  = BuildAttributeCheckbox(LoadStringResource(nullptr, IDS_CHANGE_ATTR_LABEL_READONLY), _options.readOnly);
        _systemCheckbox    = BuildAttributeCheckbox(LoadStringResource(nullptr, IDS_CHANGE_ATTR_LABEL_SYSTEM), _options.system);
        _encryptedCheckbox = _root->AddChild<Checkbox>(LoadStringResource(nullptr, IDS_CHANGE_ATTR_LABEL_ENCRYPTED));
        _encryptedCheckbox->SetEnabled(false);

        _dateTimeSectionLabel = _root->AddChild<Label>(LoadStringResource(nullptr, IDS_CHANGE_ATTR_SECTION_DATETIME));
        _dateTimeSectionLabel->SetAlignment(DWRITE_TEXT_ALIGNMENT_LEADING);
        _dateTimeSectionLabel->SetFontRole(FontRole::BodyStrong);

        BuildDateTimeRow(_modifiedTimeRow, LoadStringResource(nullptr, IDS_CHANGE_ATTR_LABEL_MODIFIED), _options.modifiedTime);
        BuildDateTimeRow(_createdTimeRow, LoadStringResource(nullptr, IDS_CHANGE_ATTR_LABEL_CREATED), _options.createdTime);
        BuildDateTimeRow(_accessedTimeRow, LoadStringResource(nullptr, IDS_CHANGE_ATTR_LABEL_ACCESSED), _options.accessedTime);

        _setCurrentButton = _root->AddChild<Button>(LoadStringResource(nullptr, IDS_CHANGE_ATTR_SET_CURRENT));
        _setCurrentButton->SetOnClick([this] { SetAllDateTimeRowsToCurrent(); });

        _optionsSectionLabel = _root->AddChild<Label>(LoadStringResource(nullptr, IDS_CHANGE_ATTR_SECTION_OPTIONS));
        _optionsSectionLabel->SetAlignment(DWRITE_TEXT_ALIGNMENT_LEADING);
        _optionsSectionLabel->SetFontRole(FontRole::BodyStrong);

        _includeSubdirectoriesCheckbox = _root->AddChild<Checkbox>(LoadStringResource(nullptr, IDS_CHANGE_ATTR_INCLUDE_SUBDIRECTORIES));
        _includeSubdirectoriesCheckbox->SetEnabled(_allowSubdirectories);
        _includeSubdirectoriesCheckbox->SetChecked(_allowSubdirectories && _options.includeSubdirectories);

        _removeStreamsCheckbox = _root->AddChild<Checkbox>(LoadStringResource(nullptr, IDS_CHANGE_ATTR_REMOVE_STREAMS));
        _removeStreamsCheckbox->SetChecked(_options.removeAlternateDataStreams);

        _okButton = _root->AddChild<Button>(LoadStringResource(nullptr, IDS_BTN_OK));
        _okButton->SetPrimary(true);
        _okButton->SetOnClick([this] { Confirm(); });

        _cancelButton = _root->AddChild<Button>(LoadStringResource(nullptr, IDS_BTN_CANCEL));
        _cancelButton->SetOnClick([this] { Cancel(); });

        _dxHost.SetRoot(std::move(_rootStorage));
    }

    RedSalamander::DxUi::Checkbox* BuildAttributeCheckbox(std::wstring text, FolderWindow::AttributeChangeState state)
    {
        using namespace RedSalamander::DxUi;

        Checkbox* checkbox = _root->AddChild<Checkbox>(std::move(text));
        ApplyAttributeCheckboxState(checkbox, state);
        checkbox->SetOnToggled([this, checkbox](bool) noexcept { CycleAttributeCheckbox(checkbox); });
        return checkbox;
    }

    void BuildDateTimeRow(DateTimeRowControls& row, std::wstring label, const FolderWindow::ChangeAttributesOptions::TimestampOption& option)
    {
        using namespace RedSalamander::DxUi;

        const int64_t ticks = option.value != 0 ? option.value : GetCurrentFileTimeTicksForChangeAttributes();

        row.label = _root->AddChild<Label>(std::move(label));
        row.label->SetAlignment(DWRITE_TEXT_ALIGNMENT_LEADING);
        row.checkbox = _root->AddChild<Checkbox>(L"");
        row.checkbox->SetChecked(option.enabled);
        row.date = _root->AddChild<TextField>(FormatDateForChangeAttributes(ticks));
        row.time = _root->AddChild<TextField>(FormatTimeForChangeAttributes(ticks));

        const auto enableRow = [checkbox = row.checkbox](std::wstring_view) noexcept
        {
            if (checkbox && ! checkbox->IsChecked())
            {
                checkbox->SetChecked(true);
            }
        };
        row.date->SetOnTextChanged(enableRow);
        row.time->SetOnTextChanged(enableRow);
    }

    void CycleAttributeCheckbox(RedSalamander::DxUi::Checkbox* checkbox) noexcept
    {
        if (! checkbox)
        {
            return;
        }

        FolderWindow::AttributeChangeState* state = nullptr;
        if (checkbox == _readOnlyCheckbox)
        {
            state = &_options.readOnly;
        }
        else if (checkbox == _hiddenCheckbox)
        {
            state = &_options.hidden;
        }
        else if (checkbox == _systemCheckbox)
        {
            state = &_options.system;
        }
        else if (checkbox == _archiveCheckbox)
        {
            state = &_options.archive;
        }

        if (! state)
        {
            return;
        }

        *state = NextAttributeChangeState(*state);
        ApplyAttributeCheckboxState(checkbox, *state);
    }

    void SetDateTimeRowToTicks(DateTimeRowControls& row, int64_t ticks, bool checked) noexcept
    {
        if (row.checkbox)
        {
            row.checkbox->SetChecked(checked);
        }
        if (row.date)
        {
            row.date->SetText(FormatDateForChangeAttributes(ticks));
        }
        if (row.time)
        {
            row.time->SetText(FormatTimeForChangeAttributes(ticks));
        }
    }

    void SetAllDateTimeRowsToCurrent() noexcept
    {
        const int64_t now = GetCurrentFileTimeTicksForChangeAttributes();
        SetDateTimeRowToTicks(_modifiedTimeRow, now, true);
        SetDateTimeRowToTicks(_createdTimeRow, now, true);
        SetDateTimeRowToTicks(_accessedTimeRow, now, true);
    }

    void ApplyTheme() noexcept
    {
        _palette = MakeAppThemeDxPalette(_theme, _theme.windowBackground);
        _dxHost.SetTheme(_palette);
        if (_hWnd)
        {
            ApplyWindowChromeTheme(_hWnd.get(), _theme, WindowBackdropTarget::Tool, GetActiveWindow() == _hWnd.get());
            static_cast<void>(RedrawWindow(_hWnd.get(), nullptr, nullptr, RDW_FRAME | RDW_INVALIDATE | RDW_UPDATENOW));
        }
    }

    void Layout() noexcept
    {
        if (! _root)
        {
            return;
        }

        const auto startedAt     = std::chrono::steady_clock::now();
        const D2D1_RECT_F client = _dxHost.GetClientBoundsDip();
        _root->SetBounds(client);

        constexpr float kMarginDip            = 20.0f;
        constexpr float kGapDip               = 8.0f;
        constexpr float kSectionHeightDip     = 22.0f;
        constexpr float kCheckboxHeightDip    = 28.0f;
        constexpr float kDateTimeRowHeightDip = 32.0f;
        constexpr float kSetCurrentWidthDip   = 116.0f;
        constexpr float kButtonWidthDip       = 110.0f;
        constexpr float kButtonHeightDip      = 34.0f;

        const float left        = client.left + kMarginDip;
        const float right       = std::max(left, client.right - kMarginDip);
        const float columnWidth = std::max(96.0f, (right - left) / 3.0f);
        float y                 = client.top + kMarginDip;

        if (_attributesSectionLabel)
        {
            _attributesSectionLabel->SetBounds(D2D1::RectF(left, y, right, y + kSectionHeightDip));
        }
        y += kSectionHeightDip + (kGapDip * 0.5f);

        LayoutCheckboxGridRow(_archiveCheckbox, _hiddenCheckbox, _compressedCheckbox, left, columnWidth, right, y, kCheckboxHeightDip);
        y += kCheckboxHeightDip;
        LayoutCheckboxGridRow(_readOnlyCheckbox, _systemCheckbox, _encryptedCheckbox, left, columnWidth, right, y, kCheckboxHeightDip);
        y += kCheckboxHeightDip + (kGapDip * 2.0f);

        if (_dateTimeSectionLabel)
        {
            _dateTimeSectionLabel->SetBounds(D2D1::RectF(left, y, right, y + kSectionHeightDip));
        }
        y += kSectionHeightDip + (kGapDip * 0.5f);

        LayoutDateTimeRow(_modifiedTimeRow, left, right, y, kDateTimeRowHeightDip);
        y += kDateTimeRowHeightDip;
        LayoutDateTimeRow(_createdTimeRow, left, right, y, kDateTimeRowHeightDip);
        y += kDateTimeRowHeightDip;
        LayoutDateTimeRow(_accessedTimeRow, left, right, y, kDateTimeRowHeightDip);
        y += kDateTimeRowHeightDip + (kGapDip * 0.5f);

        if (_setCurrentButton)
        {
            const float buttonLeft = std::max(left, right - kSetCurrentWidthDip);
            _setCurrentButton->SetBounds(D2D1::RectF(buttonLeft, y, buttonLeft + kSetCurrentWidthDip, y + kButtonHeightDip));
        }
        y += kButtonHeightDip + (kGapDip * 1.5f);

        if (_optionsSectionLabel)
        {
            _optionsSectionLabel->SetBounds(D2D1::RectF(left, y, right, y + kSectionHeightDip));
        }
        y += kSectionHeightDip + (kGapDip * 0.5f);

        if (_includeSubdirectoriesCheckbox)
        {
            _includeSubdirectoriesCheckbox->SetBounds(D2D1::RectF(left, y, right, y + kCheckboxHeightDip));
        }
        y += kCheckboxHeightDip;

        if (_removeStreamsCheckbox)
        {
            _removeStreamsCheckbox->SetBounds(D2D1::RectF(left, y, right, y + kCheckboxHeightDip));
        }

        const float buttonsTop = std::max(y + kCheckboxHeightDip + kGapDip, client.bottom - kMarginDip - kButtonHeightDip);
        const float cancelLeft = std::max(left, right - kButtonWidthDip);
        const float okLeft     = std::max(left, cancelLeft - kGapDip - kButtonWidthDip);

        if (_okButton)
        {
            _okButton->SetBounds(D2D1::RectF(okLeft, buttonsTop, okLeft + kButtonWidthDip, buttonsTop + kButtonHeightDip));
        }
        if (_cancelButton)
        {
            _cancelButton->SetBounds(D2D1::RectF(cancelLeft, buttonsTop, cancelLeft + kButtonWidthDip, buttonsTop + kButtonHeightDip));
        }

        Debug::Perf::Emit(L"commands.dialog.changeAttributes_layout_us",
                          L"",
                          Debug::Perf::ElapsedUs(startedAt),
                          15u,
                          static_cast<uint64_t>(CountVisibleChildWindowsLocal(_hWnd.get())),
                          S_OK);
    }

    static void LayoutCheckboxGridRow(RedSalamander::DxUi::Checkbox* first,
                                      RedSalamander::DxUi::Checkbox* second,
                                      RedSalamander::DxUi::Checkbox* third,
                                      float left,
                                      float columnWidth,
                                      float right,
                                      float top,
                                      float height) noexcept
    {
        const auto layout = [&](RedSalamander::DxUi::Checkbox* checkbox, float column) noexcept
        {
            if (! checkbox)
            {
                return;
            }

            const float columnLeft  = left + (column * columnWidth);
            const float columnRight = std::min(right, columnLeft + columnWidth - 6.0f);
            checkbox->SetBounds(D2D1::RectF(columnLeft, top, columnRight, top + height));
        };

        layout(first, 0.0f);
        layout(second, 1.0f);
        layout(third, 2.0f);
    }

    static void LayoutDateTimeRow(const DateTimeRowControls& row, float left, float right, float top, float height) noexcept
    {
        constexpr float kLabelWidthDip = 92.0f;
        constexpr float kCheckSizeDip  = 28.0f;
        constexpr float kDateWidthDip  = 126.0f;
        constexpr float kTimeWidthDip  = 96.0f;
        constexpr float kGapDip        = 8.0f;

        const float fieldHeight = std::min(28.0f, height);
        const float fieldTop    = top + std::max(0.0f, (height - fieldHeight) * 0.5f);
        float x                 = left;

        if (row.label)
        {
            row.label->SetBounds(D2D1::RectF(x, top, x + kLabelWidthDip, top + height));
        }
        x += kLabelWidthDip;

        if (row.checkbox)
        {
            row.checkbox->SetBounds(D2D1::RectF(x, top, x + kCheckSizeDip, top + height));
        }
        x += kCheckSizeDip + kGapDip;

        const float dateRight = std::min(right, x + kDateWidthDip);
        if (row.date)
        {
            row.date->SetBounds(D2D1::RectF(x, fieldTop, dateRight, fieldTop + fieldHeight));
        }
        x = dateRight + kGapDip;

        const float timeRight = std::min(right, x + kTimeWidthDip);
        if (row.time)
        {
            row.time->SetBounds(D2D1::RectF(x, fieldTop, timeRight, fieldTop + fieldHeight));
        }
    }

    [[nodiscard]] bool ReadDateTimeRow(const DateTimeRowControls& row, FolderWindow::ChangeAttributesOptions::TimestampOption& option) const noexcept
    {
        option.enabled = row.checkbox && row.checkbox->IsChecked();
        if (! option.enabled)
        {
            return true;
        }

        int64_t parsedTicks = 0;
        if (! row.date || ! row.time || ! TryParseChangeAttributesDateTime(row.date->GetText(), row.time->GetText(), parsedTicks))
        {
            return false;
        }

        option.value = parsedTicks;
        return true;
    }

    [[nodiscard]] bool ReadOptionsFromUi(FolderWindow::ChangeAttributesOptions& options) const noexcept
    {
        options.readOnly                   = _options.readOnly;
        options.hidden                     = _options.hidden;
        options.system                     = _options.system;
        options.archive                    = _options.archive;
        options.includeSubdirectories      = _allowSubdirectories && _includeSubdirectoriesCheckbox && _includeSubdirectoriesCheckbox->IsChecked();
        options.removeAlternateDataStreams = _removeStreamsCheckbox && _removeStreamsCheckbox->IsChecked();
        return ReadDateTimeRow(_modifiedTimeRow, options.modifiedTime) && ReadDateTimeRow(_createdTimeRow, options.createdTime) &&
               ReadDateTimeRow(_accessedTimeRow, options.accessedTime);
    }

    void Confirm() noexcept
    {
        FolderWindow::ChangeAttributesOptions options{};
        if (! ReadOptionsFromUi(options))
        {
            MessageBeep(MB_ICONWARNING);
            return;
        }

        _options = options;
        _result  = _options;
        _done    = true;
        if (_hWnd && IsWindow(_hWnd.get()) != FALSE)
        {
            _hWnd.reset();
        }
    }

    void Cancel() noexcept
    {
        _result.reset();
        _done = true;
        if (_hWnd && IsWindow(_hWnd.get()) != FALSE)
        {
            _hWnd.reset();
        }
    }

    void SetState(FolderWindow::AttributeChangeState readOnly,
                  FolderWindow::AttributeChangeState hidden,
                  FolderWindow::AttributeChangeState system,
                  FolderWindow::AttributeChangeState archive,
                  bool removeAlternateDataStreams) noexcept
    {
        _options.readOnly = readOnly;
        _options.hidden   = hidden;
        _options.system   = system;
        _options.archive  = archive;
        ApplyAttributeCheckboxState(_readOnlyCheckbox, readOnly);
        ApplyAttributeCheckboxState(_hiddenCheckbox, hidden);
        ApplyAttributeCheckboxState(_systemCheckbox, system);
        ApplyAttributeCheckboxState(_archiveCheckbox, archive);
        if (_removeStreamsCheckbox)
        {
            _removeStreamsCheckbox->SetChecked(removeAlternateDataStreams);
        }
    }

#ifdef ENABLE_TESTS
    LRESULT OnDebugCommand(ChangeAttributesOptionsPromptDebugCommand command, LPARAM lParam) noexcept
    {
        switch (command)
        {
            case ChangeAttributesOptionsPromptDebugCommand::GetSnapshot:
            {
                auto* snapshot = reinterpret_cast<ChangeAttributesOptionsPromptDebugSnapshot*>(lParam);
                if (! snapshot)
                {
                    return FALSE;
                }

                FolderWindow::ChangeAttributesOptions options{};
                static_cast<void>(ReadOptionsFromUi(options));
                *snapshot                                = ChangeAttributesOptionsPromptDebugSnapshot{};
                snapshot->usesDxUiHost                   = _dxHost.GetRoot() != nullptr;
                snapshot->visibleChildWindowCount        = CountVisibleChildWindowsLocal(_hWnd.get());
                snapshot->visibleNativeChildControlCount = CountVisibleNativeChildControlWindowsLocal(_hWnd.get());
                snapshot->dialogClassName                = GetWindowClassNameLocal(_hWnd.get());
                snapshot->readOnly                       = static_cast<uint8_t>(options.readOnly);
                snapshot->hidden                         = static_cast<uint8_t>(options.hidden);
                snapshot->system                         = static_cast<uint8_t>(options.system);
                snapshot->archive                        = static_cast<uint8_t>(options.archive);
                snapshot->dateTimeSectionVisible         = _dateTimeSectionLabel && _dateTimeSectionLabel->IsVisible();
                snapshot->modifiedTimeVisible            = _modifiedTimeRow.checkbox && _modifiedTimeRow.checkbox->IsVisible();
                snapshot->createdTimeVisible             = _createdTimeRow.checkbox && _createdTimeRow.checkbox->IsVisible();
                snapshot->accessedTimeVisible            = _accessedTimeRow.checkbox && _accessedTimeRow.checkbox->IsVisible();
                snapshot->includeSubdirectoriesVisible   = _includeSubdirectoriesCheckbox && _includeSubdirectoriesCheckbox->IsVisible();
                snapshot->includeSubdirectoriesEnabled   = _includeSubdirectoriesCheckbox && _includeSubdirectoriesCheckbox->IsEnabled();
                snapshot->includeSubdirectoriesChecked   = options.includeSubdirectories;
                snapshot->removeAlternateDataStreams     = options.removeAlternateDataStreams;
                return TRUE;
            }
            case ChangeAttributesOptionsPromptDebugCommand::SetState:
            {
                const auto* payload = reinterpret_cast<const ChangeAttributesOptionsPromptDebugStatePayload*>(lParam);
                if (! payload)
                {
                    return FALSE;
                }

                SetState(AttributeChangeStateFromRaw(payload->readOnly),
                         AttributeChangeStateFromRaw(payload->hidden),
                         AttributeChangeStateFromRaw(payload->system),
                         AttributeChangeStateFromRaw(payload->archive),
                         payload->removeAlternateDataStreams);
                return TRUE;
            }
            case ChangeAttributesOptionsPromptDebugCommand::CycleArchive: CycleAttributeCheckbox(_archiveCheckbox); return TRUE;
            case ChangeAttributesOptionsPromptDebugCommand::Confirm: Confirm(); return TRUE;
            case ChangeAttributesOptionsPromptDebugCommand::Cancel: Cancel(); return TRUE;
        }

        return FALSE;
    }
#endif

private:
    HWND _ownerWindow        = nullptr;
    HWND _restoreFocusWindow = nullptr;
    AppTheme _theme{};
    FolderWindow::ChangeAttributesOptions _options{};
    bool _allowSubdirectories = false;
    wil::unique_hwnd _hWnd;
    RedSalamander::DxUi::WindowHost _dxHost;
    RedSalamander::DxUi::ThemePalette _palette{};
    std::unique_ptr<RedSalamander::DxUi::Panel> _rootStorage;
    RedSalamander::DxUi::Panel* _root                   = nullptr;
    RedSalamander::DxUi::Label* _attributesSectionLabel = nullptr;
    RedSalamander::DxUi::Label* _dateTimeSectionLabel   = nullptr;
    RedSalamander::DxUi::Label* _optionsSectionLabel    = nullptr;
    RedSalamander::DxUi::Checkbox* _archiveCheckbox     = nullptr;
    RedSalamander::DxUi::Checkbox* _hiddenCheckbox      = nullptr;
    RedSalamander::DxUi::Checkbox* _compressedCheckbox  = nullptr;
    RedSalamander::DxUi::Checkbox* _readOnlyCheckbox    = nullptr;
    RedSalamander::DxUi::Checkbox* _systemCheckbox      = nullptr;
    RedSalamander::DxUi::Checkbox* _encryptedCheckbox   = nullptr;
    DateTimeRowControls _modifiedTimeRow;
    DateTimeRowControls _createdTimeRow;
    DateTimeRowControls _accessedTimeRow;
    RedSalamander::DxUi::Button* _setCurrentButton                = nullptr;
    RedSalamander::DxUi::Checkbox* _includeSubdirectoriesCheckbox = nullptr;
    RedSalamander::DxUi::Checkbox* _removeStreamsCheckbox         = nullptr;
    RedSalamander::DxUi::Button* _okButton                        = nullptr;
    RedSalamander::DxUi::Button* _cancelButton                    = nullptr;
    bool _done                                                    = false;
    std::optional<FolderWindow::ChangeAttributesOptions> _result;
};

[[nodiscard]] std::optional<FolderWindow::ChangeAttributesOptions> PromptForChangeAttributes(HWND ownerWindow,
                                                                                             const AppTheme& theme,
                                                                                             FolderWindow::ChangeAttributesOptions initialOptions,
                                                                                             bool allowSubdirectories) noexcept
{
    ChangeAttributesOptionsPromptWindow prompt(ownerWindow, theme, initialOptions, allowSubdirectories);
    return prompt.ShowModal();
}

[[nodiscard]] bool IsChangeAttributesNoop(const FolderWindow::ChangeAttributesOptions& options) noexcept
{
    return options.readOnly == FolderWindow::AttributeChangeState::LeaveUnchanged && options.hidden == FolderWindow::AttributeChangeState::LeaveUnchanged &&
           options.system == FolderWindow::AttributeChangeState::LeaveUnchanged && options.archive == FolderWindow::AttributeChangeState::LeaveUnchanged &&
           ! options.modifiedTime.enabled && ! options.createdTime.enabled && ! options.accessedTime.enabled && ! options.removeAlternateDataStreams;
}

[[nodiscard]] unsigned long ApplyAttributeChange(unsigned long attributes, FolderWindow::AttributeChangeState state, unsigned long attributeFlag) noexcept
{
    switch (state)
    {
        case FolderWindow::AttributeChangeState::Set: return attributes | attributeFlag;
        case FolderWindow::AttributeChangeState::Clear: return attributes & ~attributeFlag;
        case FolderWindow::AttributeChangeState::LeaveUnchanged: return attributes;
    }

    return attributes;
}

[[nodiscard]] unsigned long ResolveDesiredAttributes(unsigned long attributes, const FolderWindow::ChangeAttributesOptions& options) noexcept
{
    unsigned long desired = attributes;
    desired               = ApplyAttributeChange(desired, options.readOnly, FILE_ATTRIBUTE_READONLY);
    desired               = ApplyAttributeChange(desired, options.hidden, FILE_ATTRIBUTE_HIDDEN);
    desired               = ApplyAttributeChange(desired, options.system, FILE_ATTRIBUTE_SYSTEM);
    desired               = ApplyAttributeChange(desired, options.archive, FILE_ATTRIBUTE_ARCHIVE);
    return desired;
}

void ApplyTimestampOptionsForChangeAttributes(FileSystemBasicInformation& info, const FolderWindow::ChangeAttributesOptions& options) noexcept
{
    if (options.modifiedTime.enabled)
    {
        info.lastWriteTime = options.modifiedTime.value;
    }
    if (options.createdTime.enabled)
    {
        info.creationTime = options.createdTime.value;
    }
    if (options.accessedTime.enabled)
    {
        info.lastAccessTime = options.accessedTime.value;
    }
}

[[nodiscard]] bool ChangeAttributesBasicInfoEquals(const FileSystemBasicInformation& left, const FileSystemBasicInformation& right) noexcept
{
    return left.creationTime == right.creationTime && left.lastAccessTime == right.lastAccessTime && left.lastWriteTime == right.lastWriteTime &&
           left.attributes == right.attributes;
}

[[nodiscard]] bool ChangeAttributesTimesEqual(const FileSystemBasicInformation& left, const FileSystemBasicInformation& right) noexcept
{
    return left.creationTime == right.creationTime && left.lastAccessTime == right.lastAccessTime && left.lastWriteTime == right.lastWriteTime;
}

[[nodiscard]] HRESULT SetChangeAttributesBasicInfo(IFileSystemIO* fileIo,
                                                   const std::filesystem::path& path,
                                                   FileSystemBasicInformation& info,
                                                   const FileSystemBasicInformation& desired) noexcept
{
    if (! fileIo)
    {
        return E_POINTER;
    }

    if (ChangeAttributesBasicInfoEquals(info, desired))
    {
        return S_FALSE;
    }

    FileSystemBasicInformation updated = desired;
    updated.sizeBytes                  = sizeof(FileSystemBasicInformation);
    const HRESULT hr                   = fileIo->SetFileBasicInformation(path.c_str(), &updated);
    if (SUCCEEDED(hr))
    {
        info = updated;
    }
    return hr;
}

[[nodiscard]] std::wstring BuildChangeAttributesReportSummary(const FolderWindow::ChangeAttributesReport& report)
{
    if (report.failures == 0u)
    {
        return FormatStringResource(
            nullptr, IDS_FMT_CHANGE_ATTRIBUTES_REPORT, report.itemsProcessed, report.attributesChanged, report.timesChanged, report.streamsRemoved);
    }

    return FormatStringResource(nullptr,
                                IDS_FMT_CHANGE_ATTRIBUTES_REPORT_FAILURE,
                                report.itemsProcessed,
                                report.attributesChanged,
                                report.timesChanged,
                                report.streamsRemoved,
                                report.failures,
                                static_cast<unsigned long>(static_cast<uint32_t>(report.firstFailure)));
}

struct ChangeAttributesTaskPayload final
{
    FolderWindow::InformationalTaskUpdate update{};
};

struct ChangeAttributesCompletedPayload final
{
    FolderWindow::Pane pane = FolderWindow::Pane::Left;
    HRESULT hr              = S_OK;
    FolderWindow::ChangeAttributesReport report{};
    bool refreshNeeded = false;
};

[[nodiscard]] bool IsChangeAttributesDirectory(unsigned long attributes) noexcept
{
    return (attributes & FILE_ATTRIBUTE_DIRECTORY) != 0;
}

[[nodiscard]] bool ShouldEnumerateChangeAttributesDirectory(unsigned long attributes) noexcept
{
    return IsChangeAttributesDirectory(attributes) && (attributes & FILE_ATTRIBUTE_REPARSE_POINT) == 0;
}

[[nodiscard]] bool ChangeAttributesSelectionContainsDirectory(IFileSystemIO* fileIo, const std::vector<std::filesystem::path>& paths) noexcept
{
    if (! fileIo)
    {
        return false;
    }

    for (const std::filesystem::path& path : paths)
    {
        FileSystemBasicInformation info{};
        info.sizeBytes = sizeof(FileSystemBasicInformation);
        if (SUCCEEDED(fileIo->GetFileBasicInformation(path.c_str(), &info)) && IsChangeAttributesDirectory(info.attributes))
        {
            return true;
        }
    }

    return false;
}

[[nodiscard]] FolderWindow::ChangeAttributesOptions BuildInitialChangeAttributesOptions(IFileSystemIO* fileIo,
                                                                                        const std::vector<std::filesystem::path>& paths) noexcept
{
    FolderWindow::ChangeAttributesOptions options{};
    int64_t timestamp = GetCurrentFileTimeTicksForChangeAttributes();

    if (fileIo && ! paths.empty())
    {
        FileSystemBasicInformation info{};
        info.sizeBytes = sizeof(FileSystemBasicInformation);
        if (SUCCEEDED(fileIo->GetFileBasicInformation(paths.front().c_str(), &info)))
        {
            options.modifiedTime.value = info.lastWriteTime != 0 ? info.lastWriteTime : timestamp;
            options.createdTime.value  = info.creationTime != 0 ? info.creationTime : timestamp;
            options.accessedTime.value = info.lastAccessTime != 0 ? info.lastAccessTime : timestamp;
            return options;
        }
    }

    options.modifiedTime.value = timestamp;
    options.createdTime.value  = timestamp;
    options.accessedTime.value = timestamp;
    return options;
}

void ApplyChangeAttributesToPath(IFileSystemIO* fileIo,
                                 IFileSystemItemStreams* streamOps,
                                 const std::filesystem::path& path,
                                 const FolderWindow::ChangeAttributesOptions& options,
                                 FolderWindow::ChangeAttributesReport& report,
                                 bool& refreshNeeded) noexcept
{
    ++report.itemsProcessed;

    FileSystemBasicInformation info{};
    info.sizeBytes      = sizeof(FileSystemBasicInformation);
    const HRESULT getHr = fileIo ? fileIo->GetFileBasicInformation(path.c_str(), &info) : E_POINTER;
    if (FAILED(getHr))
    {
        RecordChangeAttributesFailure(report, getHr);
        return;
    }

    const FileSystemBasicInformation original = info;
    FileSystemBasicInformation desired        = info;
    desired.attributes                        = ResolveDesiredAttributes(original.attributes, options);
    ApplyTimestampOptionsForChangeAttributes(desired, options);

    const bool attributesWillChange = desired.attributes != original.attributes;
    const bool timesWillChange      = ! ChangeAttributesTimesEqual(desired, original);

    if (options.removeAlternateDataStreams)
    {
        FileSystemBasicInformation streamReady = info;
        streamReady.attributes                 = desired.attributes;
        if ((streamReady.attributes & FILE_ATTRIBUTE_READONLY) != 0)
        {
            streamReady.attributes &= ~FILE_ATTRIBUTE_READONLY;
        }

        const HRESULT streamReadyHr = SetChangeAttributesBasicInfo(fileIo, path, info, streamReady);
        if (FAILED(streamReadyHr))
        {
            RecordChangeAttributesFailure(report, streamReadyHr);
        }

        std::vector<std::wstring> streamNames;
        const HRESULT collectHr = CollectRemovableStreamsForChangeAttributes(fileIo, path, streamNames);
        if (FAILED(collectHr))
        {
            RecordChangeAttributesFailure(report, collectHr);
        }
        else if (! streamNames.empty())
        {
            if (! streamOps)
            {
                RecordChangeAttributesFailure(report, HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED));
            }
            else
            {
                for (const std::wstring& streamName : streamNames)
                {
                    const auto streamStartedAt = std::chrono::steady_clock::now();
                    const HRESULT deleteHr     = streamOps->DeleteItemStream(path.c_str(), streamName.c_str());
                    Debug::Perf::Emit(
                        L"fileattrs.stream_remove_us", streamName, Debug::Perf::ElapsedUs(streamStartedAt), 1u, SUCCEEDED(deleteHr) ? 1u : 0u, deleteHr);
                    if (FAILED(deleteHr))
                    {
                        RecordChangeAttributesFailure(report, deleteHr);
                    }
                    else
                    {
                        ++report.streamsRemoved;
                        refreshNeeded = true;
                    }
                }
            }
        }
    }

    const HRESULT finalHr = SetChangeAttributesBasicInfo(fileIo, path, info, desired);
    if (FAILED(finalHr))
    {
        RecordChangeAttributesFailure(report, finalHr);
        return;
    }

    if (attributesWillChange && info.attributes == desired.attributes)
    {
        ++report.attributesChanged;
        refreshNeeded = true;
    }
    if (timesWillChange && ChangeAttributesTimesEqual(info, desired))
    {
        ++report.timesChanged;
        refreshNeeded = true;
    }
}

struct ChangeAttributesProgressState final
{
    HWND hwnd               = nullptr;
    FolderWindow::Pane pane = FolderWindow::Pane::Left;
    std::wstring title;
    uint64_t taskId          = 0;
    ULONGLONG lastPostedTick = 0;

    std::filesystem::path currentPath;
    uint64_t scannedFolders = 0;
    uint64_t scannedEntries = 0;
    uint64_t plannedItems   = 0;
    uint64_t completedItems = 0;
    bool enumerating        = false;
    bool applying           = false;

    void PostTaskUpdate(bool finished, HRESULT hr, std::wstring doneSummary = {}) noexcept
    {
        if (! hwnd || IsWindow(hwnd) == FALSE || taskId == 0)
        {
            return;
        }

        FolderWindow::InformationalTaskUpdate info{};
        info.kind                           = FolderWindow::InformationalTaskUpdate::Kind::ChangeAttributes;
        info.taskId                         = taskId;
        info.title                          = title;
        info.changeAttributesCurrentPath    = currentPath;
        info.changeAttributesScannedFolders = scannedFolders;
        info.changeAttributesScannedEntries = scannedEntries;
        info.changeAttributesPlannedItems   = plannedItems;
        info.changeAttributesCompletedItems = completedItems;
        info.changeAttributesEnumerating    = ! finished && enumerating;
        info.changeAttributesApplying       = ! finished && applying;
        info.finished                       = finished;
        info.resultHr                       = hr;
        info.doneSummary                    = std::move(doneSummary);

        auto payload    = std::make_unique<ChangeAttributesTaskPayload>();
        payload->update = std::move(info);
        static_cast<void>(PostMessagePayload(hwnd, WndMsg::kChangeAttributesTaskUpdate, 0, std::move(payload)));
    }

    void MaybePostTaskUpdate() noexcept
    {
        const ULONGLONG nowTick = GetTickCount64();
        if (lastPostedTick != 0 && nowTick >= lastPostedTick && (nowTick - lastPostedTick) < 100ull)
        {
            return;
        }

        lastPostedTick = nowTick;
        PostTaskUpdate(false, S_OK);
    }
};

[[nodiscard]] HRESULT CollectRecursiveChangeAttributesPaths(IFileSystem* fileSystem,
                                                            IFileSystemIO* fileIo,
                                                            const std::vector<std::filesystem::path>& roots,
                                                            const std::stop_token& stopToken,
                                                            ChangeAttributesProgressState& progress,
                                                            FolderWindow::ChangeAttributesReport& report,
                                                            std::vector<std::filesystem::path>& outPaths) noexcept
{
    if (! fileSystem || ! fileIo)
    {
        return E_POINTER;
    }

    outPaths.clear();
    std::vector<std::filesystem::path> pendingFolders;

    progress.enumerating = true;
    progress.applying    = false;

    for (const std::filesystem::path& root : roots)
    {
        if (stopToken.stop_requested())
        {
            return HRESULT_FROM_WIN32(ERROR_CANCELLED);
        }

        outPaths.push_back(root);
        progress.currentPath  = root;
        progress.plannedItems = static_cast<uint64_t>(outPaths.size());
        progress.MaybePostTaskUpdate();

        FileSystemBasicInformation info{};
        info.sizeBytes = sizeof(FileSystemBasicInformation);
        if (SUCCEEDED(fileIo->GetFileBasicInformation(root.c_str(), &info)) && ShouldEnumerateChangeAttributesDirectory(info.attributes))
        {
            pendingFolders.push_back(root);
        }
    }

    while (! pendingFolders.empty())
    {
        if (stopToken.stop_requested())
        {
            return HRESULT_FROM_WIN32(ERROR_CANCELLED);
        }

        const std::filesystem::path folder = pendingFolders.back();
        pendingFolders.pop_back();

        progress.currentPath = folder;
        ++progress.scannedFolders;
        progress.MaybePostTaskUpdate();

        wil::com_ptr<IFilesInformation> filesInformation;
        const HRESULT readHr = fileSystem->ReadDirectoryInfo(folder.c_str(), filesInformation.put());
        if (FAILED(readHr))
        {
            RecordChangeAttributesFailure(report, readHr);
            continue;
        }

        FileInfo* entry        = nullptr;
        const HRESULT bufferHr = filesInformation ? filesInformation->GetBuffer(&entry) : E_POINTER;
        if (FAILED(bufferHr))
        {
            RecordChangeAttributesFailure(report, bufferHr);
            continue;
        }

        while (entry != nullptr)
        {
            if (stopToken.stop_requested())
            {
                return HRESULT_FROM_WIN32(ERROR_CANCELLED);
            }

            const size_t nameChars = static_cast<size_t>(entry->FileNameSize) / sizeof(wchar_t);
            const std::wstring_view name(entry->FileName, nameChars);
            if (name != L"." && name != L"..")
            {
                const std::filesystem::path child = folder / std::wstring(name);
                outPaths.push_back(child);
                ++progress.scannedEntries;
                progress.currentPath  = child;
                progress.plannedItems = static_cast<uint64_t>(outPaths.size());
                if (ShouldEnumerateChangeAttributesDirectory(entry->FileAttributes))
                {
                    pendingFolders.push_back(child);
                }
                progress.MaybePostTaskUpdate();
            }

            if (entry->NextEntryOffset == 0)
            {
                break;
            }

            entry = reinterpret_cast<FileInfo*>(reinterpret_cast<std::byte*>(entry) + entry->NextEntryOffset);
        }
    }

    return S_OK;
}

void ShowChangeAttributesOverlay(
    FolderWindow& window, FolderWindow::Pane pane, FolderView::OverlaySeverity severity, UINT messageStringId, HRESULT hr = S_OK) noexcept
{
    Debug::Perf::Scope perf(L"fileattrs.feedback_us");
    perf.SetHr(hr);

    std::wstring title = LoadStringResource(nullptr, IDS_CMD_CHANGE_ATTRIBUTES);
    if (title.empty())
    {
        title = LoadStringResource(nullptr, severity == FolderView::OverlaySeverity::Error ? IDS_CAPTION_ERROR : IDS_CAPTION_WARNING);
    }

    std::wstring message = LoadStringResource(nullptr, messageStringId);
    if (message.empty())
    {
        message = title;
    }

    window.ShowPaneAlertOverlay(pane, FolderView::ErrorOverlayKind::Operation, severity, std::move(title), std::move(message), hr, true, false);
}

void ShowChangeAttributesReportOverlay(FolderWindow& window, FolderWindow::Pane pane, const FolderWindow::ChangeAttributesReport& report) noexcept
{
    Debug::Perf::Scope perf(L"fileattrs.feedback_us");
    perf.SetHr(report.firstFailure);

    std::wstring title = LoadStringResource(nullptr, IDS_CMD_CHANGE_ATTRIBUTES);
    if (title.empty())
    {
        if (report.failures == 0u)
        {
            title = LoadEmbeddedStringResource(nullptr, IDS_APP_TITLE);
        }
        else
        {
            title = LoadStringResource(nullptr, IDS_CAPTION_WARNING);
        }
    }

    std::wstring message = report.summary;
    if (message.empty())
    {
        message = BuildChangeAttributesReportSummary(report);
    }

    window.ShowPaneAlertOverlay(pane,
                                FolderView::ErrorOverlayKind::Operation,
                                report.failures == 0u ? FolderView::OverlaySeverity::Information : FolderView::OverlaySeverity::Warning,
                                std::move(title),
                                std::move(message),
                                report.firstFailure,
                                report.failures != 0u,
                                false);
}

using MakeFileListSettings = Common::Settings::MakeFileListSettings;

#ifdef ENABLE_TESTS
std::optional<MakeFileListSettings> g_makeFileListAutomation;
std::atomic_uint32_t g_makeFileListWorkerDelayMs{0u};
std::atomic_bool g_makeFileListWorkerActive{false};
#endif

struct MakeFileListTaskPayload final
{
    FolderWindow::InformationalTaskUpdate update{};
};

struct MakeFileListCompletedPayload final
{
    FolderWindow::Pane pane = FolderWindow::Pane::Left;
    uint64_t taskId         = 0u;
    std::wstring title;
    MakeFileListSettings options{};
    std::filesystem::path currentFolder;
    std::wstring clipboardText;
    std::wstring outputTarget;
    uint64_t entryCount      = 0u;
    uint64_t collectFailures = 0u;
    uint64_t outputBytes     = 0u;
    uint64_t outputElapsedUs = 0u;
    uint64_t totalElapsedUs  = 0u;
    HRESULT hr               = S_OK;
};

struct MakeFileListProgressState final
{
    HWND hwnd = nullptr;
    std::wstring title;
    uint64_t taskId          = 0u;
    ULONGLONG lastPostedTick = 0u;

    bool collecting = false;
    bool rendering  = false;
    bool writing    = false;
    std::filesystem::path currentPath;
    uint64_t scannedFolders  = 0u;
    uint64_t scannedEntries  = 0u;
    uint64_t totalEntries    = 0u;
    uint64_t renderedEntries = 0u;

    void PostTaskUpdate(bool finished, HRESULT hr, std::wstring doneSummary = {}) noexcept
    {
        if (! hwnd || IsWindow(hwnd) == FALSE || taskId == 0u)
        {
            return;
        }

        FolderWindow::InformationalTaskUpdate info{};
        info.kind                        = FolderWindow::InformationalTaskUpdate::Kind::MakeFileList;
        info.taskId                      = taskId;
        info.title                       = title;
        info.makeFileListCollecting      = ! finished && collecting;
        info.makeFileListRendering       = ! finished && rendering;
        info.makeFileListWriting         = ! finished && writing;
        info.makeFileListCurrentPath     = currentPath;
        info.makeFileListScannedFolders  = scannedFolders;
        info.makeFileListScannedEntries  = scannedEntries;
        info.makeFileListTotalEntries    = totalEntries;
        info.makeFileListRenderedEntries = renderedEntries;
        info.finished                    = finished;
        info.resultHr                    = hr;
        info.doneSummary                 = std::move(doneSummary);

        auto payload    = std::make_unique<MakeFileListTaskPayload>();
        payload->update = std::move(info);
        static_cast<void>(PostMessagePayload(hwnd, WndMsg::kMakeFileListTaskUpdate, 0, std::move(payload)));
    }

    void MaybePostTaskUpdate() noexcept
    {
        const ULONGLONG nowTick = GetTickCount64();
        if (lastPostedTick != 0u && nowTick >= lastPostedTick && (nowTick - lastPostedTick) < 100ull)
        {
            return;
        }

        lastPostedTick = nowTick;
        PostTaskUpdate(false, S_OK);
    }
};

struct MakeFileListEntry final
{
    std::filesystem::path path;
    std::wstring name;
    std::wstring fullPath;
    bool isDirectory = false;
    uint64_t size    = 0u;
    bool hasSize     = false;
    FILETIME modified{};
    bool hasModified   = false;
    DWORD attributes   = 0u;
    bool hasAttributes = false;
};

[[nodiscard]] std::string Utf8FromUtf16ForMakeFileList(std::wstring_view text) noexcept
{
    return Common::Strings::Utf8FromUtf16StrictOrEmpty(text);
}

[[nodiscard]] std::wstring Utf16FromUtf8ForMakeFileList(std::string_view text) noexcept
{
    return Common::Strings::Utf16FromUtf8StrictOrEmpty(text);
}

[[nodiscard]] std::wstring LowerAsciiMacroToken(std::wstring_view text)
{
    std::wstring result;
    result.reserve(text.size());
    for (const wchar_t ch : text)
    {
        result.push_back(static_cast<wchar_t>(std::towlower(static_cast<wint_t>(ch))));
    }
    return result;
}

[[nodiscard]] std::wstring FormatMakeFileListTime(const FILETIME& fileTime)
{
    SYSTEMTIME systemTime{};
    if (FileTimeToSystemTime(&fileTime, &systemTime) == FALSE)
    {
        return {};
    }

    return std::format(L"{:04}-{:02}-{:02}T{:02}:{:02}:{:02}Z",
                       systemTime.wYear,
                       systemTime.wMonth,
                       systemTime.wDay,
                       systemTime.wHour,
                       systemTime.wMinute,
                       systemTime.wSecond);
}

[[nodiscard]] std::wstring FormatMakeFileListAttributes(DWORD attributes)
{
    return std::format(L"{:08X}", static_cast<unsigned>(attributes));
}

[[nodiscard]] std::wstring MakeFileListEntryValue(const MakeFileListEntry& entry, std::wstring_view token)
{
    const std::wstring lower = LowerAsciiMacroToken(token);
    if (lower == L"name" || lower == L"filename")
    {
        return entry.name;
    }
    if (lower == L"fullpath" || lower == L"pathandfilename")
    {
        return entry.fullPath;
    }
    if (lower == L"path")
    {
        const std::filesystem::path parent = entry.path.parent_path();
        return parent.empty() ? std::wstring{} : parent.wstring();
    }
    if (lower == L"size")
    {
        return entry.hasSize ? std::format(L"{}", entry.size) : std::wstring{};
    }
    if (lower == L"modified")
    {
        return entry.hasModified ? FormatMakeFileListTime(entry.modified) : std::wstring{};
    }
    if (lower == L"attributes")
    {
        return entry.hasAttributes ? FormatMakeFileListAttributes(entry.attributes) : std::wstring{};
    }
    if (lower == L"isdirectory")
    {
        return entry.isDirectory ? std::wstring(L"true") : std::wstring(L"false");
    }

    return {};
}

[[nodiscard]] bool IsKnownMakeFileListMacro(std::wstring_view token) noexcept
{
    const std::wstring lower = LowerAsciiMacroToken(token);
    return lower == L"name" || lower == L"filename" || lower == L"fullpath" || lower == L"pathandfilename" || lower == L"path" || lower == L"size" ||
           lower == L"modified" || lower == L"attributes" || lower == L"isdirectory";
}

[[nodiscard]] std::wstring ExpandMakeFileListTextMacro(std::wstring_view macro, const MakeFileListEntry& entry)
{
    std::wstring result;
    result.reserve(macro.size() + 64u);

    for (size_t index = 0u; index < macro.size(); ++index)
    {
        const wchar_t ch = macro[index];
        if (ch == L'{' && (index + 1u) < macro.size() && macro[index + 1u] == L'{')
        {
            result.push_back(L'{');
            ++index;
            continue;
        }
        if (ch == L'}' && (index + 1u) < macro.size() && macro[index + 1u] == L'}')
        {
            result.push_back(L'}');
            ++index;
            continue;
        }
        if (ch != L'{')
        {
            result.push_back(ch);
            continue;
        }

        const size_t close = macro.find(L'}', index + 1u);
        if (close == std::wstring_view::npos)
        {
            result.push_back(ch);
            continue;
        }

        const std::wstring_view token = macro.substr(index + 1u, close - index - 1u);
        if (! IsKnownMakeFileListMacro(token))
        {
            result.append(macro.substr(index, close - index + 1u));
            index = close;
            continue;
        }

        result.append(MakeFileListEntryValue(entry, token));
        index = close;
    }

    return result;
}

enum class MakeFileListEntryReadResult : uint8_t
{
    Added,
    Skipped,
    Failed,
};

[[nodiscard]] MakeFileListEntryReadResult TryReadMakeFileListEntry(const std::filesystem::path& path, bool includeDirectories, MakeFileListEntry& out) noexcept
{
    WIN32_FILE_ATTRIBUTE_DATA data{};
    if (GetFileAttributesExW(path.c_str(), GetFileExInfoStandard, &data) == FALSE)
    {
        return MakeFileListEntryReadResult::Failed;
    }

    const bool isDirectory = (data.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY) != 0u;
    if (isDirectory && ! includeDirectories)
    {
        return MakeFileListEntryReadResult::Skipped;
    }

    out               = {};
    out.path          = path;
    out.name          = path.filename().wstring();
    out.fullPath      = path.wstring();
    out.isDirectory   = isDirectory;
    out.modified      = data.ftLastWriteTime;
    out.hasModified   = true;
    out.attributes    = data.dwFileAttributes;
    out.hasAttributes = true;

    if (! isDirectory)
    {
        out.size    = (static_cast<uint64_t>(data.nFileSizeHigh) << 32u) | static_cast<uint64_t>(data.nFileSizeLow);
        out.hasSize = true;
    }

    if (out.name.empty())
    {
        out.name = out.fullPath;
    }

    return MakeFileListEntryReadResult::Added;
}

[[nodiscard]] HRESULT AddMakeFileListPath(const std::filesystem::path& path,
                                          bool includeDirectories,
                                          const std::stop_token& stopToken,
                                          MakeFileListProgressState& progress,
                                          std::vector<MakeFileListEntry>& entries,
                                          uint64_t& failures) noexcept
{
    if (stopToken.stop_requested())
    {
        return HRESULT_FROM_WIN32(ERROR_CANCELLED);
    }

    progress.currentPath = path;
    ++progress.scannedEntries;

    MakeFileListEntry entry{};
    const MakeFileListEntryReadResult readResult = TryReadMakeFileListEntry(path, includeDirectories, entry);
    if (readResult == MakeFileListEntryReadResult::Added)
    {
        if (entry.isDirectory)
        {
            ++progress.scannedFolders;
        }
        entries.push_back(std::move(entry));
    }
    else if (readResult == MakeFileListEntryReadResult::Failed)
    {
        std::error_code ec;
        if (std::filesystem::exists(path, ec) || ec)
        {
            ++failures;
        }
    }

    progress.MaybePostTaskUpdate();
    return S_OK;
}

[[nodiscard]] HRESULT CollectMakeFileListDirectoryContents(const std::filesystem::path& root,
                                                           bool recursive,
                                                           bool includeDirectories,
                                                           const std::stop_token& stopToken,
                                                           MakeFileListProgressState& progress,
                                                           std::vector<MakeFileListEntry>& entries,
                                                           uint64_t& failures) noexcept
{
    constexpr std::filesystem::directory_options options = std::filesystem::directory_options::skip_permission_denied;
    std::error_code ec;

    if (recursive)
    {
        std::filesystem::recursive_directory_iterator it(root, options, ec);
        const std::filesystem::recursive_directory_iterator end;
        while (! ec && it != end)
        {
            if (stopToken.stop_requested())
            {
                return HRESULT_FROM_WIN32(ERROR_CANCELLED);
            }

            const std::filesystem::path path = it->path();
            if (const HRESULT hr = AddMakeFileListPath(path, includeDirectories, stopToken, progress, entries, failures); FAILED(hr))
            {
                return hr;
            }
            it.increment(ec);
        }
    }
    else
    {
        std::filesystem::directory_iterator it(root, options, ec);
        const std::filesystem::directory_iterator end;
        while (! ec && it != end)
        {
            if (stopToken.stop_requested())
            {
                return HRESULT_FROM_WIN32(ERROR_CANCELLED);
            }

            const std::filesystem::path path = it->path();
            if (const HRESULT hr = AddMakeFileListPath(path, includeDirectories, stopToken, progress, entries, failures); FAILED(hr))
            {
                return hr;
            }
            it.increment(ec);
        }
    }

    if (ec)
    {
        ++failures;
    }
    return stopToken.stop_requested() ? HRESULT_FROM_WIN32(ERROR_CANCELLED) : S_OK;
}

[[nodiscard]] HRESULT CollectMakeFileListEntries(const std::filesystem::path& currentFolder,
                                                 const std::vector<std::filesystem::path>& selectedPaths,
                                                 const MakeFileListSettings& options,
                                                 const std::stop_token& stopToken,
                                                 MakeFileListProgressState& progress,
                                                 uint64_t& failures,
                                                 std::vector<MakeFileListEntry>& entries) noexcept
{
    failures = 0u;
    entries.clear();
    progress.currentPath = currentFolder;

    if (options.sourceMode == Common::Settings::MakeFileListSourceMode::CurrentFolder)
    {
        if (const HRESULT hr =
                CollectMakeFileListDirectoryContents(currentFolder, options.recursive, options.includeDirectories, stopToken, progress, entries, failures);
            FAILED(hr))
        {
            return hr;
        }
    }
    else
    {
        entries.reserve(selectedPaths.size());
        for (const std::filesystem::path& path : selectedPaths)
        {
            if (const HRESULT hr = AddMakeFileListPath(path, options.includeDirectories, stopToken, progress, entries, failures); FAILED(hr))
            {
                return hr;
            }

            std::error_code ec;
            if (options.recursive && std::filesystem::is_directory(path, ec))
            {
                if (const HRESULT hr = CollectMakeFileListDirectoryContents(path, true, options.includeDirectories, stopToken, progress, entries, failures);
                    FAILED(hr))
                {
                    return hr;
                }
            }
        }
    }

    if (stopToken.stop_requested())
    {
        return HRESULT_FROM_WIN32(ERROR_CANCELLED);
    }

    std::sort(entries.begin(),
              entries.end(),
              [](const MakeFileListEntry& left, const MakeFileListEntry& right) noexcept
    {
        const std::wstring leftPath  = left.fullPath;
        const std::wstring rightPath = right.fullPath;
        const int compare            = CompareStringOrdinal(leftPath.c_str(), -1, rightPath.c_str(), -1, TRUE);
        if (compare == CSTR_EQUAL)
        {
            return leftPath < rightPath;
        }
        return compare == CSTR_LESS_THAN;
    });

    progress.totalEntries = static_cast<uint64_t>(entries.size());
    return S_OK;
}

[[nodiscard]] HRESULT AddMakeFileListJsonString(yyjson_mut_doc* doc, yyjson_mut_val* obj, const char* key, std::wstring_view value) noexcept
{
    const std::string utf8 = Utf8FromUtf16ForMakeFileList(value);
    if (! value.empty() && utf8.empty())
    {
        return HRESULT_FROM_WIN32(ERROR_NO_UNICODE_TRANSLATION);
    }

    if (! yyjson_mut_obj_add_strncpy(doc, obj, key, utf8.data(), utf8.size()))
    {
        return E_OUTOFMEMORY;
    }

    return S_OK;
}

[[nodiscard]] HRESULT AddMakeFileListJsonEntry(yyjson_mut_doc* doc,
                                               yyjson_mut_val* entries,
                                               const MakeFileListEntry& entry,
                                               const MakeFileListSettings& options) noexcept
{
    yyjson_mut_val* item = yyjson_mut_obj(doc);
    if (! item)
    {
        return E_OUTOFMEMORY;
    }

    if (options.includeName)
    {
        if (const HRESULT hr = AddMakeFileListJsonString(doc, item, "name", entry.name); FAILED(hr))
        {
            return hr;
        }
    }
    if (options.includeFullPath)
    {
        if (const HRESULT hr = AddMakeFileListJsonString(doc, item, "fullPath", entry.fullPath); FAILED(hr))
        {
            return hr;
        }
    }
    if (options.includeSize)
    {
        if (entry.hasSize)
        {
            if (! yyjson_mut_obj_add_uint(doc, item, "size", entry.size))
            {
                return E_OUTOFMEMORY;
            }
        }
        else
        {
            yyjson_mut_val* nullValue = yyjson_mut_null(doc);
            if (! nullValue || ! yyjson_mut_obj_add_val(doc, item, "size", nullValue))
            {
                return E_OUTOFMEMORY;
            }
        }
    }
    if (options.includeModified)
    {
        if (const HRESULT hr = AddMakeFileListJsonString(doc, item, "modified", entry.hasModified ? FormatMakeFileListTime(entry.modified) : std::wstring{});
            FAILED(hr))
        {
            return hr;
        }
    }
    if (options.includeAttributes)
    {
        if (const HRESULT hr =
                AddMakeFileListJsonString(doc, item, "attributes", entry.hasAttributes ? FormatMakeFileListAttributes(entry.attributes) : std::wstring{});
            FAILED(hr))
        {
            return hr;
        }
    }

    if (! yyjson_mut_obj_add_bool(doc, item, "directory", entry.isDirectory) || ! yyjson_mut_arr_add_val(entries, item))
    {
        return E_OUTOFMEMORY;
    }

    return S_OK;
}

[[nodiscard]] HRESULT RenderMakeFileListJson(const std::vector<MakeFileListEntry>& entries,
                                             const MakeFileListSettings& options,
                                             const std::stop_token& stopToken,
                                             MakeFileListProgressState& progress,
                                             std::string& outUtf8) noexcept
{
    outUtf8.clear();

    yyjson_mut_doc* doc = yyjson_mut_doc_new(nullptr);
    if (! doc)
    {
        return E_OUTOFMEMORY;
    }
    const auto freeDoc = wil::scope_exit([&] { yyjson_mut_doc_free(doc); });

    yyjson_mut_val* root = yyjson_mut_obj(doc);
    if (! root)
    {
        return E_OUTOFMEMORY;
    }
    yyjson_mut_doc_set_root(doc, root);

    if (! yyjson_mut_obj_add_str(doc, root, "format", "json") || ! yyjson_mut_obj_add_uint(doc, root, "count", entries.size()))
    {
        return E_OUTOFMEMORY;
    }

    yyjson_mut_val* array = yyjson_mut_arr(doc);
    if (! array || ! yyjson_mut_obj_add_val(doc, root, "entries", array))
    {
        return E_OUTOFMEMORY;
    }

    for (const MakeFileListEntry& entry : entries)
    {
        if (stopToken.stop_requested())
        {
            return HRESULT_FROM_WIN32(ERROR_CANCELLED);
        }
        if (const HRESULT hr = AddMakeFileListJsonEntry(doc, array, entry, options); FAILED(hr))
        {
            return hr;
        }
        ++progress.renderedEntries;
        progress.MaybePostTaskUpdate();
    }

    if (stopToken.stop_requested())
    {
        return HRESULT_FROM_WIN32(ERROR_CANCELLED);
    }

    yyjson_write_err writeErr{};
    size_t jsonLen = 0u;
    char* json     = yyjson_mut_write_opts(doc, YYJSON_WRITE_NOFLAG, nullptr, &jsonLen, &writeErr);
    if (! json)
    {
        return E_OUTOFMEMORY;
    }
    const auto freeJson = wil::scope_exit([&] { std::free(json); });

    outUtf8.assign(json, jsonLen);
    return S_OK;
}

void AppendMakeFileListCsvField(std::wstring& output, std::wstring_view value)
{
    const bool quote = value.find_first_of(L"\",\r\n") != std::wstring_view::npos;
    if (quote)
    {
        output.push_back(L'"');
    }

    for (const wchar_t ch : value)
    {
        if (ch == L'"')
        {
            output.append(L"\"\"");
        }
        else
        {
            output.push_back(ch);
        }
    }

    if (quote)
    {
        output.push_back(L'"');
    }
}

void AppendMakeFileListCsvRow(std::wstring& output, const std::vector<std::wstring>& fields)
{
    for (size_t index = 0u; index < fields.size(); ++index)
    {
        if (index != 0u)
        {
            output.push_back(L',');
        }
        AppendMakeFileListCsvField(output, fields[index]);
    }
    output.append(L"\r\n");
}

[[nodiscard]] std::vector<std::wstring> MakeFileListCsvFields(const MakeFileListEntry& entry, const MakeFileListSettings& options)
{
    std::vector<std::wstring> fields;
    fields.reserve(5u);
    if (options.includeName)
    {
        fields.push_back(entry.name);
    }
    if (options.includeFullPath)
    {
        fields.push_back(entry.fullPath);
    }
    if (options.includeSize)
    {
        fields.push_back(entry.hasSize ? std::format(L"{}", entry.size) : std::wstring{});
    }
    if (options.includeModified)
    {
        fields.push_back(entry.hasModified ? FormatMakeFileListTime(entry.modified) : std::wstring{});
    }
    if (options.includeAttributes)
    {
        fields.push_back(entry.hasAttributes ? FormatMakeFileListAttributes(entry.attributes) : std::wstring{});
    }
    return fields;
}

[[nodiscard]] std::vector<std::wstring> MakeFileListCsvHeader(const MakeFileListSettings& options)
{
    std::vector<std::wstring> fields;
    fields.reserve(5u);
    if (options.includeName)
    {
        fields.emplace_back(L"name");
    }
    if (options.includeFullPath)
    {
        fields.emplace_back(L"fullPath");
    }
    if (options.includeSize)
    {
        fields.emplace_back(L"size");
    }
    if (options.includeModified)
    {
        fields.emplace_back(L"modified");
    }
    if (options.includeAttributes)
    {
        fields.emplace_back(L"attributes");
    }
    return fields;
}

[[nodiscard]] HRESULT RenderMakeFileListCsvWide(const std::vector<MakeFileListEntry>& entries,
                                                const MakeFileListSettings& options,
                                                const std::stop_token& stopToken,
                                                MakeFileListProgressState& progress,
                                                std::wstring& output) noexcept
{
    output.clear();
    output.reserve(entries.size() * 96u);
    AppendMakeFileListCsvRow(output, MakeFileListCsvHeader(options));
    for (const MakeFileListEntry& entry : entries)
    {
        if (stopToken.stop_requested())
        {
            return HRESULT_FROM_WIN32(ERROR_CANCELLED);
        }
        AppendMakeFileListCsvRow(output, MakeFileListCsvFields(entry, options));
        ++progress.renderedEntries;
        progress.MaybePostTaskUpdate();
    }
    return S_OK;
}

[[nodiscard]] HRESULT RenderMakeFileListTextWide(const std::vector<MakeFileListEntry>& entries,
                                                 const MakeFileListSettings& options,
                                                 const std::stop_token& stopToken,
                                                 MakeFileListProgressState& progress,
                                                 std::wstring& output) noexcept
{
    const std::wstring_view macro = options.textMacro.empty() ? std::wstring_view(L"{fullPath}") : std::wstring_view(options.textMacro);

    output.clear();
    output.reserve(entries.size() * (macro.size() + 32u));
    for (const MakeFileListEntry& entry : entries)
    {
        if (stopToken.stop_requested())
        {
            return HRESULT_FROM_WIN32(ERROR_CANCELLED);
        }
        output.append(ExpandMakeFileListTextMacro(macro, entry));
        output.append(L"\r\n");
        ++progress.renderedEntries;
        progress.MaybePostTaskUpdate();
    }
    return S_OK;
}

[[nodiscard]] HRESULT RenderMakeFileListOutput(const std::vector<MakeFileListEntry>& entries,
                                               const MakeFileListSettings& options,
                                               const std::stop_token& stopToken,
                                               MakeFileListProgressState& progress,
                                               std::string& outUtf8,
                                               std::wstring& outClipboardText) noexcept
{
    outUtf8.clear();
    outClipboardText.clear();
    progress.renderedEntries = 0u;

    if (stopToken.stop_requested())
    {
        return HRESULT_FROM_WIN32(ERROR_CANCELLED);
    }

    if (options.format == Common::Settings::MakeFileListFormat::Json)
    {
        if (const HRESULT hr = RenderMakeFileListJson(entries, options, stopToken, progress, outUtf8); FAILED(hr))
        {
            return hr;
        }
        outClipboardText = Utf16FromUtf8ForMakeFileList(outUtf8);
        if (stopToken.stop_requested())
        {
            return HRESULT_FROM_WIN32(ERROR_CANCELLED);
        }
        return (! outUtf8.empty() && outClipboardText.empty()) ? HRESULT_FROM_WIN32(ERROR_NO_UNICODE_TRANSLATION) : S_OK;
    }

    const HRESULT renderHr = options.format == Common::Settings::MakeFileListFormat::Csv
                                 ? RenderMakeFileListCsvWide(entries, options, stopToken, progress, outClipboardText)
                                 : RenderMakeFileListTextWide(entries, options, stopToken, progress, outClipboardText);
    if (FAILED(renderHr))
    {
        return renderHr;
    }
    outUtf8 = Utf8FromUtf16ForMakeFileList(outClipboardText);
    if (stopToken.stop_requested())
    {
        return HRESULT_FROM_WIN32(ERROR_CANCELLED);
    }
    return (! outClipboardText.empty() && outUtf8.empty()) ? HRESULT_FROM_WIN32(ERROR_NO_UNICODE_TRANSLATION) : S_OK;
}

[[nodiscard]] HRESULT WriteMakeFileListUtf8File(const std::filesystem::path& path, std::string_view bytes, const std::stop_token& stopToken) noexcept
{
    if (stopToken.stop_requested())
    {
        return HRESULT_FROM_WIN32(ERROR_CANCELLED);
    }

    Common::Files::LocalFileTransaction transaction;
    HRESULT hr = Common::Files::LocalFileTransaction::Create(path, Common::Files::ExistingTargetPolicy::Replace, true, transaction);
    if (FAILED(hr))
    {
        return hr;
    }

    hr = transaction.Write(bytes);
    if (FAILED(hr))
    {
        return hr;
    }
    if (stopToken.stop_requested())
    {
        return HRESULT_FROM_WIN32(ERROR_CANCELLED);
    }
    return transaction.Commit(static_cast<uint64_t>(bytes.size()));
}

[[nodiscard]] const wchar_t* MakeFileListFormatDetail(Common::Settings::MakeFileListFormat format) noexcept
{
    switch (format)
    {
        case Common::Settings::MakeFileListFormat::Json: return L"json";
        case Common::Settings::MakeFileListFormat::Csv: return L"csv";
        case Common::Settings::MakeFileListFormat::Text: return L"text";
    }
    return L"text";
}

[[nodiscard]] const wchar_t* MakeFileListDefaultExtension(Common::Settings::MakeFileListFormat format) noexcept
{
    switch (format)
    {
        case Common::Settings::MakeFileListFormat::Json: return L"json";
        case Common::Settings::MakeFileListFormat::Csv: return L"csv";
        case Common::Settings::MakeFileListFormat::Text: return L"txt";
    }
    return L"txt";
}

[[nodiscard]] DWORD MakeFileListFilterIndex(Common::Settings::MakeFileListFormat format) noexcept
{
    switch (format)
    {
        case Common::Settings::MakeFileListFormat::Text: return 1u;
        case Common::Settings::MakeFileListFormat::Csv: return 2u;
        case Common::Settings::MakeFileListFormat::Json: return 3u;
    }
    return 1u;
}

[[nodiscard]] std::optional<std::filesystem::path> PromptForMakeFileListOutputFile(HWND owner,
                                                                                   const std::filesystem::path& currentFolder,
                                                                                   const MakeFileListSettings& options) noexcept
{
    constexpr size_t kFileBufferChars = 32768u;
    std::vector<wchar_t> fileBuffer(kFileBufferChars, L'\0');

    std::filesystem::path suggested = options.outputFile;
    if (suggested.empty())
    {
        suggested = currentFolder / std::format(L"file-list.{}", MakeFileListDefaultExtension(options.format));
    }
    else if (suggested.is_relative())
    {
        suggested = currentFolder / suggested;
    }

    const std::wstring suggestedText = suggested.wstring();
    if (suggestedText.size() >= fileBuffer.size())
    {
        return std::nullopt;
    }
    std::copy(suggestedText.begin(), suggestedText.end(), fileBuffer.begin());

    const std::wstring filter     = LoadStringResource(nullptr, IDS_MAKE_FILE_LIST_FILE_FILTER);
    const std::wstring initialDir = currentFolder.wstring();

    OPENFILENAMEW ofn{};
    ofn.lStructSize     = sizeof(ofn);
    ofn.hwndOwner       = owner;
    ofn.lpstrFilter     = filter.c_str();
    ofn.nFilterIndex    = MakeFileListFilterIndex(options.format);
    ofn.lpstrFile       = fileBuffer.data();
    ofn.nMaxFile        = static_cast<DWORD>(fileBuffer.size());
    ofn.lpstrInitialDir = initialDir.empty() ? nullptr : initialDir.c_str();
    ofn.lpstrDefExt     = MakeFileListDefaultExtension(options.format);
    ofn.Flags           = OFN_NOCHANGEDIR | OFN_HIDEREADONLY | OFN_OVERWRITEPROMPT | OFN_PATHMUSTEXIST;

    if (GetSaveFileNameW(&ofn) == FALSE)
    {
        return std::nullopt;
    }

    std::filesystem::path selected(fileBuffer.data());
    return selected.empty() ? std::nullopt : std::optional<std::filesystem::path>{std::move(selected)};
}

void ShowMakeFileListOverlay(
    FolderWindow& window, FolderWindow::Pane pane, FolderView::OverlaySeverity severity, std::wstring message, HRESULT hr = S_OK) noexcept
{
    Debug::Perf::Scope perf(L"makeFileList.feedback_us");
    perf.SetHr(hr);

    std::wstring title = LoadStringResource(nullptr, IDS_CMD_MAKE_FILE_LIST);
    if (title.empty())
    {
        title = LoadStringResource(nullptr, severity == FolderView::OverlaySeverity::Error ? IDS_CAPTION_ERROR : IDS_CAPTION_WARNING);
    }

    window.ShowPaneAlertOverlay(pane,
                                FolderView::ErrorOverlayKind::Operation,
                                severity,
                                std::move(title),
                                std::move(message),
                                hr,
                                severity != FolderView::OverlaySeverity::Information,
                                false);
}

constexpr wchar_t kMakeFileListOptionsPromptClassName[] = L"RedSalamander.MakeFileListOptionsPrompt";

#ifdef ENABLE_TESTS
std::atomic<HWND> g_makeFileListOptionsPromptWindow{nullptr};

enum class MakeFileListOptionsPromptDebugCommand : uintptr_t
{
    GetSnapshot = 1u,
    SetState,
    Confirm,
    Cancel,
};

struct MakeFileListOptionsPromptDebugStatePayload final
{
    uint8_t sourceMode   = 0u;
    bool recursive       = false;
    uint8_t format       = 0u;
    uint8_t outputTarget = 0u;
    std::wstring textMacro;
    std::wstring outputFile;
    bool includeName        = false;
    bool includeFullPath    = false;
    bool includeSize        = false;
    bool includeModified    = false;
    bool includeAttributes  = false;
    bool includeDirectories = false;
};

[[nodiscard]] UINT GetMakeFileListOptionsPromptDebugMessage() noexcept
{
    static const UINT message = RegisterWindowMessageW(L"RedSalamander.MakeFileListOptionsPrompt.Debug");
    return message;
}
#endif

[[nodiscard]] std::optional<size_t> MakeFileListSourceModeToIndex(Common::Settings::MakeFileListSourceMode sourceMode) noexcept
{
    return sourceMode == Common::Settings::MakeFileListSourceMode::CurrentFolder ? std::optional<size_t>{1u} : std::optional<size_t>{0u};
}

[[nodiscard]] Common::Settings::MakeFileListSourceMode MakeFileListSourceModeFromIndex(std::optional<size_t> index) noexcept
{
    return index.value_or(0u) == 1u ? Common::Settings::MakeFileListSourceMode::CurrentFolder : Common::Settings::MakeFileListSourceMode::Selection;
}

[[nodiscard]] std::optional<size_t> MakeFileListFormatToIndex(Common::Settings::MakeFileListFormat format) noexcept
{
    switch (format)
    {
        case Common::Settings::MakeFileListFormat::Csv: return 1u;
        case Common::Settings::MakeFileListFormat::Json: return 2u;
        case Common::Settings::MakeFileListFormat::Text: return 0u;
    }
    return 0u;
}

[[nodiscard]] Common::Settings::MakeFileListFormat MakeFileListFormatFromIndex(std::optional<size_t> index) noexcept
{
    switch (index.value_or(0u))
    {
        case 1u: return Common::Settings::MakeFileListFormat::Csv;
        case 2u: return Common::Settings::MakeFileListFormat::Json;
        default: return Common::Settings::MakeFileListFormat::Text;
    }
}

[[nodiscard]] std::optional<size_t> MakeFileListOutputTargetToIndex(Common::Settings::MakeFileListOutputTarget outputTarget) noexcept
{
    return outputTarget == Common::Settings::MakeFileListOutputTarget::File ? std::optional<size_t>{1u} : std::optional<size_t>{0u};
}

[[nodiscard]] Common::Settings::MakeFileListOutputTarget MakeFileListOutputTargetFromIndex(std::optional<size_t> index) noexcept
{
    return index.value_or(0u) == 1u ? Common::Settings::MakeFileListOutputTarget::File : Common::Settings::MakeFileListOutputTarget::Clipboard;
}

#ifdef ENABLE_TESTS
[[nodiscard]] Common::Settings::MakeFileListSourceMode MakeFileListSourceModeFromRaw(uint8_t value) noexcept
{
    return value == static_cast<uint8_t>(Common::Settings::MakeFileListSourceMode::CurrentFolder) ? Common::Settings::MakeFileListSourceMode::CurrentFolder
                                                                                                  : Common::Settings::MakeFileListSourceMode::Selection;
}

[[nodiscard]] Common::Settings::MakeFileListFormat MakeFileListFormatFromRaw(uint8_t value) noexcept
{
    switch (value)
    {
        case static_cast<uint8_t>(Common::Settings::MakeFileListFormat::Csv): return Common::Settings::MakeFileListFormat::Csv;
        case static_cast<uint8_t>(Common::Settings::MakeFileListFormat::Json): return Common::Settings::MakeFileListFormat::Json;
        default: return Common::Settings::MakeFileListFormat::Text;
    }
}

[[nodiscard]] Common::Settings::MakeFileListOutputTarget MakeFileListOutputTargetFromRaw(uint8_t value) noexcept
{
    return value == static_cast<uint8_t>(Common::Settings::MakeFileListOutputTarget::File) ? Common::Settings::MakeFileListOutputTarget::File
                                                                                           : Common::Settings::MakeFileListOutputTarget::Clipboard;
}
#endif

[[nodiscard]] std::vector<RedSalamander::DxUi::ComboBox::Item> BuildMakeFileListSourceItems()
{
    using RedSalamander::DxUi::ComboBox;
    return {
        ComboBox::Item{L"selection", LoadStringResource(nullptr, IDS_MAKE_FILE_LIST_SOURCE_SELECTION)},
        ComboBox::Item{L"currentFolder", LoadStringResource(nullptr, IDS_MAKE_FILE_LIST_SOURCE_CURRENT_FOLDER)},
    };
}

[[nodiscard]] std::vector<RedSalamander::DxUi::ComboBox::Item> BuildMakeFileListFormatItems()
{
    using RedSalamander::DxUi::ComboBox;
    const std::wstring textFormat = LoadStringResource(nullptr, IDS_MAKE_FILE_LIST_FORMAT_TEXT);
    const std::wstring csvFormat  = LoadEmbeddedStringResource(nullptr, IDS_MAKE_FILE_LIST_FORMAT_CSV);
    const std::wstring jsonFormat = LoadEmbeddedStringResource(nullptr, IDS_MAKE_FILE_LIST_FORMAT_JSON);
    return {
        ComboBox::Item{L"text", textFormat},
        ComboBox::Item{L"csv", csvFormat},
        ComboBox::Item{L"json", jsonFormat},
    };
}

[[nodiscard]] std::vector<RedSalamander::DxUi::ComboBox::Item> BuildMakeFileListOutputTargetItems()
{
    using RedSalamander::DxUi::ComboBox;
    return {
        ComboBox::Item{L"clipboard", LoadStringResource(nullptr, IDS_MAKE_FILE_LIST_OUTPUT_CLIPBOARD)},
        ComboBox::Item{L"file", LoadStringResource(nullptr, IDS_MAKE_FILE_LIST_OUTPUT_FILE)},
    };
}

class MakeFileListOptionsPromptWindow final
{
public:
    MakeFileListOptionsPromptWindow(const MakeFileListOptionsPromptWindow&)            = delete;
    MakeFileListOptionsPromptWindow& operator=(const MakeFileListOptionsPromptWindow&) = delete;
    MakeFileListOptionsPromptWindow(MakeFileListOptionsPromptWindow&&)                 = delete;
    MakeFileListOptionsPromptWindow& operator=(MakeFileListOptionsPromptWindow&&)      = delete;

    MakeFileListOptionsPromptWindow(HWND ownerWindow, const AppTheme& theme, MakeFileListSettings initialOptions) noexcept
        : _ownerWindow(GetOwnerWindowOrSelf(ownerWindow)),
          _restoreFocusWindow(ownerWindow && IsWindow(ownerWindow) != FALSE ? ownerWindow : nullptr),
          _theme(theme),
          _options(std::move(initialOptions))
    {
        if (_ownerWindow && IsWindow(_ownerWindow) != FALSE)
        {
            const HWND focused = GetFocus();
            if (focused && IsWindow(focused) != FALSE && (focused == _ownerWindow || IsChild(_ownerWindow, focused) != FALSE))
            {
                _restoreFocusWindow = focused;
            }
            else if (! _restoreFocusWindow || IsWindow(_restoreFocusWindow) == FALSE ||
                     (_restoreFocusWindow != _ownerWindow && IsChild(_ownerWindow, _restoreFocusWindow) == FALSE))
            {
                _restoreFocusWindow = _ownerWindow;
            }
        }
        else if (! _restoreFocusWindow || IsWindow(_restoreFocusWindow) == FALSE)
        {
            _restoreFocusWindow = nullptr;
        }
    }

    [[nodiscard]] std::optional<MakeFileListSettings> ShowModal() noexcept
    {
        const HRESULT classHr = EnsureWindowClass();
        if (FAILED(classHr))
        {
            return std::nullopt;
        }

        const DWORD style   = WS_CAPTION | WS_SYSMENU | WS_POPUP | WS_CLIPCHILDREN | WS_CLIPSIBLINGS;
        const DWORD exStyle = WS_EX_DLGMODALFRAME;
        const UINT dpi      = _ownerWindow && IsWindow(_ownerWindow) != FALSE ? GetDpiForWindow(_ownerWindow) : GetDpiForSystem();

        RECT bounds{0, 0, ScalePanePromptForDpi(dpi, 640), ScalePanePromptForDpi(dpi, 542)};
        if (AdjustWindowRectExForDpi(&bounds, style, FALSE, exStyle, dpi) == FALSE)
        {
            return std::nullopt;
        }

        const bool restoreOwnerEnabled = _ownerWindow && IsWindow(_ownerWindow) != FALSE && IsWindowEnabled(_ownerWindow) != FALSE;
        if (restoreOwnerEnabled)
        {
            EnableWindow(_ownerWindow, FALSE);
        }
        const auto restoreOwner = wil::scope_exit([this, restoreOwnerEnabled]() noexcept
        {
            if (restoreOwnerEnabled && _ownerWindow && IsWindow(_ownerWindow) != FALSE)
            {
                EnableWindow(_ownerWindow, TRUE);
                SetActiveWindow(_ownerWindow);
                const HWND restoreFocus = (_restoreFocusWindow && IsWindow(_restoreFocusWindow) != FALSE &&
                                           (_restoreFocusWindow == _ownerWindow || IsChild(_ownerWindow, _restoreFocusWindow) != FALSE))
                                              ? _restoreFocusWindow
                                              : _ownerWindow;
                SetFocus(restoreFocus);
            }
        });

        const std::wstring caption = LoadStringResource(nullptr, IDS_CMD_MAKE_FILE_LIST);
        const HWND hwnd            = CreateWindowExW(exStyle,
                                                     kMakeFileListOptionsPromptClassName,
                                                     caption.c_str(),
                                                     style,
                                                     CW_USEDEFAULT,
                                                     CW_USEDEFAULT,
                                                     bounds.right - bounds.left,
                                                     bounds.bottom - bounds.top,
                                                     _ownerWindow,
                                                     nullptr,
                                                     GetModuleHandleW(nullptr),
                                                     this);
        if (! hwnd)
        {
            return std::nullopt;
        }
        if (! _hWnd)
        {
            _hWnd.reset(hwnd);
        }

        static_cast<void>(Common::WindowSizing::CenterExistingWindowOnOwner(_hWnd.get(), _ownerWindow));
        static_cast<void>(_dxHost.PrimeForShow());
        ShowWindow(_hWnd.get(), SW_SHOWNORMAL);
        UpdateWindow(_hWnd.get());
        SetForegroundWindow(_hWnd.get());

        MSG msg{};
        while (! _done)
        {
            const BOOL getMessageResult = GetMessageW(&msg, nullptr, 0, 0);
            if (getMessageResult == -1)
            {
                return std::nullopt;
            }
            if (getMessageResult == 0)
            {
                _done = true;
                break;
            }

            TranslateMessage(&msg);
            DispatchMessageW(&msg);
        }

        return _result;
    }

    static LRESULT CALLBACK WndProc(HWND hwnd, UINT message, WPARAM wParam, LPARAM lParam) noexcept
    {
        if (message == WM_NCCREATE)
        {
            auto* cs   = reinterpret_cast<CREATESTRUCTW*>(lParam);
            auto* self = static_cast<MakeFileListOptionsPromptWindow*>(cs ? cs->lpCreateParams : nullptr);
            if (! self)
            {
                return FALSE;
            }

            SetWindowLongPtrW(hwnd, GWLP_USERDATA, reinterpret_cast<LONG_PTR>(self));
            if (! self->_hWnd)
            {
                self->_hWnd.reset(hwnd);
            }
            return TRUE;
        }

        auto* self = reinterpret_cast<MakeFileListOptionsPromptWindow*>(GetWindowLongPtrW(hwnd, GWLP_USERDATA));
        if (! self)
        {
            return DefWindowProcW(hwnd, message, wParam, lParam);
        }

        bool handled     = false;
        LRESULT dxResult = 0;
        if (message != WM_CREATE)
        {
            dxResult = self->_dxHost.HandleMessage(hwnd, message, wParam, lParam, handled);
        }
        if (handled)
        {
            if (message == WM_SIZE || message == WM_DPICHANGED)
            {
                self->Layout();
            }
            if (message == WM_NCDESTROY)
            {
#ifdef ENABLE_TESTS
                g_makeFileListOptionsPromptWindow.store(nullptr);
#endif
                SetWindowLongPtrW(hwnd, GWLP_USERDATA, 0);
                self->_dxHost.Detach();
                if (self->_hWnd.get() == hwnd)
                {
                    static_cast<void>(self->_hWnd.release());
                }
                self->_done = true;
            }
            return dxResult;
        }

#ifdef ENABLE_TESTS
        if (message == GetMakeFileListOptionsPromptDebugMessage())
        {
            return self->OnDebugCommand(static_cast<MakeFileListOptionsPromptDebugCommand>(wParam), lParam);
        }
#endif

        switch (message)
        {
            case WM_CREATE: return self->OnCreate(hwnd) ? 0 : -1;
            case WM_SIZE: self->Layout(); return 0;
            case WM_DPICHANGED:
            {
                const auto* suggested = reinterpret_cast<const RECT*>(lParam);
                if (suggested)
                {
                    SetWindowPos(hwnd,
                                 nullptr,
                                 suggested->left,
                                 suggested->top,
                                 suggested->right - suggested->left,
                                 suggested->bottom - suggested->top,
                                 SWP_NOZORDER | SWP_NOACTIVATE);
                }
                self->Layout();
                return 0;
            }
            case WM_ERASEBKGND: return 1;
            case WM_NCACTIVATE: ApplyTitleBarTheme(hwnd, self->_theme, wParam != FALSE); return DefWindowProcW(hwnd, message, wParam, lParam);
            case WM_CLOSE: self->Cancel(); return 0;
            case WM_NCDESTROY:
#ifdef ENABLE_TESTS
                g_makeFileListOptionsPromptWindow.store(nullptr);
#endif
                self->_dxHost.Detach();
                if (self->_hWnd.get() == hwnd)
                {
                    self->_hWnd.release();
                }
                SetWindowLongPtrW(hwnd, GWLP_USERDATA, 0);
                self->_done = true;
                break;
        }

        return DefWindowProcW(hwnd, message, wParam, lParam);
    }

private:
    [[nodiscard]] static HRESULT EnsureWindowClass() noexcept
    {
        static ATOM atom = 0;
        if (atom != 0)
        {
            return S_OK;
        }

        WNDCLASSEXW wc{};
        wc.cbSize        = sizeof(wc);
        wc.lpfnWndProc   = MakeFileListOptionsPromptWindow::WndProc;
        wc.hInstance     = GetModuleHandleW(nullptr);
        wc.hCursor       = LoadCursorW(nullptr, IDC_ARROW);
        wc.lpszClassName = kMakeFileListOptionsPromptClassName;
        wc.style         = CS_DBLCLKS;

        atom = RegisterClassExW(&wc);
        return atom != 0 ? S_OK : HRESULT_FROM_WIN32(GetLastError());
    }

    [[nodiscard]] bool OnCreate(HWND hwnd) noexcept
    {
        if (! _dxHost.Attach(hwnd))
        {
            return false;
        }

#ifdef ENABLE_TESTS
        g_makeFileListOptionsPromptWindow.store(hwnd);
#endif
        BuildUi();
        ApplyTheme();
        UpdateOutputEnabledState();
        Layout();
        if (_sourceCombo)
        {
            _dxHost.SetFocusControl(_sourceCombo);
        }
        _dxHost.SetDefaultButton(_okButton);
        _dxHost.SetCancelButton(_cancelButton);
        return true;
    }

    void BuildUi()
    {
        if (_root)
        {
            return;
        }

        using namespace RedSalamander::DxUi;

        _rootStorage = std::make_unique<Panel>();
        _root        = _rootStorage.get();

        _sourceLabel = _root->AddChild<Label>(LoadStringResource(nullptr, IDS_MAKE_FILE_LIST_SOURCE));
        _sourceCombo = _root->AddChild<ComboBox>();
        _sourceCombo->SetVariant(ComboBoxVariant::Window);
        _sourceCombo->SetEditable(false);
        _sourceCombo->SetItems(BuildMakeFileListSourceItems());
        _sourceCombo->SetSelectedIndex(MakeFileListSourceModeToIndex(_options.sourceMode));
        _sourceCombo->SetOnSubmitted([this] { Confirm(); });

        _recursiveCheckbox = _root->AddChild<Checkbox>(LoadStringResource(nullptr, IDS_MAKE_FILE_LIST_RECURSIVE));
        _recursiveCheckbox->SetChecked(_options.recursive);

        _formatLabel = _root->AddChild<Label>(LoadStringResource(nullptr, IDS_MAKE_FILE_LIST_FORMAT));
        _formatCombo = _root->AddChild<ComboBox>();
        _formatCombo->SetVariant(ComboBoxVariant::Window);
        _formatCombo->SetEditable(false);
        _formatCombo->SetItems(BuildMakeFileListFormatItems());
        _formatCombo->SetSelectedIndex(MakeFileListFormatToIndex(_options.format));
        _formatCombo->SetOnSubmitted([this] { Confirm(); });

        _textMacroLabel = _root->AddChild<Label>(LoadStringResource(nullptr, IDS_MAKE_FILE_LIST_TEXT_MACRO));
        _textMacroField = _root->AddChild<TextField>(_options.textMacro);
        _textMacroField->SetMultiline(false);
        _textMacroField->SetOnSubmitted([this] { Confirm(); });

        _fieldsLabel                = _root->AddChild<Label>(LoadStringResource(nullptr, IDS_MAKE_FILE_LIST_FIELDS));
        _includeNameCheckbox        = BuildFieldCheckbox(IDS_MAKE_FILE_LIST_INCLUDE_NAME, _options.includeName);
        _includeFullPathCheckbox    = BuildFieldCheckbox(IDS_MAKE_FILE_LIST_INCLUDE_FULLPATH, _options.includeFullPath);
        _includeSizeCheckbox        = BuildFieldCheckbox(IDS_MAKE_FILE_LIST_INCLUDE_SIZE, _options.includeSize);
        _includeModifiedCheckbox    = BuildFieldCheckbox(IDS_MAKE_FILE_LIST_INCLUDE_MODIFIED, _options.includeModified);
        _includeAttributesCheckbox  = BuildFieldCheckbox(IDS_MAKE_FILE_LIST_INCLUDE_ATTRIBUTES, _options.includeAttributes);
        _includeDirectoriesCheckbox = BuildFieldCheckbox(IDS_MAKE_FILE_LIST_INCLUDE_DIRECTORIES, _options.includeDirectories);

        _outputLabel = _root->AddChild<Label>(LoadStringResource(nullptr, IDS_MAKE_FILE_LIST_OUTPUT));
        _outputCombo = _root->AddChild<ComboBox>();
        _outputCombo->SetVariant(ComboBoxVariant::Window);
        _outputCombo->SetEditable(false);
        _outputCombo->SetItems(BuildMakeFileListOutputTargetItems());
        _outputCombo->SetSelectedIndex(MakeFileListOutputTargetToIndex(_options.outputTarget));
        _outputCombo->SetOnSelectionChanged([this](size_t) noexcept { UpdateOutputEnabledState(); });
        _outputCombo->SetOnSubmitted([this] { Confirm(); });

        _outputFileLabel = _root->AddChild<Label>(LoadStringResource(nullptr, IDS_MAKE_FILE_LIST_OUTPUT_FILE_PATH));
        _outputFileField = _root->AddChild<TextField>(_options.outputFile.wstring());
        _outputFileField->SetMultiline(false);
        _outputFileField->SetOnSubmitted([this] { Confirm(); });

        _okButton = _root->AddChild<Button>(LoadStringResource(nullptr, IDS_BTN_OK));
        _okButton->SetPrimary(true);
        _okButton->SetOnClick([this] { Confirm(); });

        _cancelButton = _root->AddChild<Button>(LoadStringResource(nullptr, IDS_BTN_CANCEL));
        _cancelButton->SetOnClick([this] { Cancel(); });

        _dxHost.SetRoot(std::move(_rootStorage));
    }

    RedSalamander::DxUi::Checkbox* BuildFieldCheckbox(UINT textId, bool checked)
    {
        RedSalamander::DxUi::Checkbox* checkbox = _root->AddChild<RedSalamander::DxUi::Checkbox>(LoadStringResource(nullptr, textId));
        checkbox->SetChecked(checked);
        return checkbox;
    }

    void ApplyTheme() noexcept
    {
        _palette = MakeAppThemeDxPalette(_theme, _theme.windowBackground);
        _dxHost.SetTheme(_palette);
        if (_hWnd)
        {
            ApplyWindowChromeTheme(_hWnd.get(), _theme, WindowBackdropTarget::Tool, GetActiveWindow() == _hWnd.get());
            static_cast<void>(RedrawWindow(_hWnd.get(), nullptr, nullptr, RDW_FRAME | RDW_INVALIDATE | RDW_UPDATENOW));
        }
    }

    void UpdateOutputEnabledState() noexcept
    {
        const bool outputToFile = MakeFileListOutputTargetFromIndex(_outputCombo ? _outputCombo->GetSelectedIndex() : std::nullopt) ==
                                  Common::Settings::MakeFileListOutputTarget::File;
        if (_outputFileField)
        {
            _outputFileField->SetEnabled(outputToFile);
        }
        if (_outputFileLabel)
        {
            _outputFileLabel->SetEnabled(outputToFile);
        }
    }

    void Layout() noexcept
    {
        if (! _root)
        {
            return;
        }

        const auto startedAt     = std::chrono::steady_clock::now();
        const D2D1_RECT_F client = _dxHost.GetClientBoundsDip();
        _root->SetBounds(client);

        constexpr float kMarginDip         = 16.0f;
        constexpr float kGapDip            = 8.0f;
        constexpr float kLabelWidthDip     = 130.0f;
        constexpr float kLabelHeightDip    = 22.0f;
        constexpr float kRowHeightDip      = 34.0f;
        constexpr float kCheckboxHeightDip = 28.0f;
        constexpr float kButtonWidthDip    = 96.0f;
        constexpr float kButtonHeightDip   = 34.0f;

        const float left      = client.left + kMarginDip;
        const float right     = std::max(left, client.right - kMarginDip);
        const float fieldLeft = std::min(right, left + kLabelWidthDip + kGapDip);
        const float mid       = left + ((right - left) * 0.5f);
        float y               = client.top + kMarginDip;

        LayoutFormRow(_sourceLabel, _sourceCombo, left, fieldLeft, right, y, kRowHeightDip);
        y += kRowHeightDip + kGapDip;

        if (_recursiveCheckbox)
        {
            _recursiveCheckbox->SetBounds(D2D1::RectF(fieldLeft, y, right, y + kCheckboxHeightDip));
        }
        y += kCheckboxHeightDip + kGapDip;

        LayoutFormRow(_formatLabel, _formatCombo, left, fieldLeft, right, y, kRowHeightDip);
        y += kRowHeightDip + kGapDip;

        LayoutFormRow(_textMacroLabel, _textMacroField, left, fieldLeft, right, y, kRowHeightDip);
        y += kRowHeightDip + (kGapDip * 1.5f);

        if (_fieldsLabel)
        {
            _fieldsLabel->SetBounds(D2D1::RectF(left, y, right, y + kLabelHeightDip));
        }
        y += kLabelHeightDip + 4.0f;

        LayoutCheckboxPair(_includeNameCheckbox, _includeFullPathCheckbox, left, mid, right, y, kCheckboxHeightDip);
        y += kCheckboxHeightDip;
        LayoutCheckboxPair(_includeSizeCheckbox, _includeModifiedCheckbox, left, mid, right, y, kCheckboxHeightDip);
        y += kCheckboxHeightDip;
        LayoutCheckboxPair(_includeAttributesCheckbox, _includeDirectoriesCheckbox, left, mid, right, y, kCheckboxHeightDip);
        y += kCheckboxHeightDip + (kGapDip * 1.5f);

        LayoutFormRow(_outputLabel, _outputCombo, left, fieldLeft, right, y, kRowHeightDip);
        y += kRowHeightDip + kGapDip;

        LayoutFormRow(_outputFileLabel, _outputFileField, left, fieldLeft, right, y, kRowHeightDip);
        y += kRowHeightDip + kGapDip;

        const float buttonsTop = std::max(y, client.bottom - kMarginDip - kButtonHeightDip);
        const float cancelLeft = std::max(left, right - kButtonWidthDip);
        const float okLeft     = std::max(left, cancelLeft - kGapDip - kButtonWidthDip);

        if (_okButton)
        {
            _okButton->SetBounds(D2D1::RectF(okLeft, buttonsTop, okLeft + kButtonWidthDip, buttonsTop + kButtonHeightDip));
        }
        if (_cancelButton)
        {
            _cancelButton->SetBounds(D2D1::RectF(cancelLeft, buttonsTop, cancelLeft + kButtonWidthDip, buttonsTop + kButtonHeightDip));
        }

        Debug::Perf::Emit(L"commands.dialog.makeFileList_layout_us",
                          L"",
                          Debug::Perf::ElapsedUs(startedAt),
                          3u,
                          static_cast<uint64_t>(CountVisibleChildWindowsLocal(_hWnd.get())),
                          S_OK);
    }

    static void LayoutFormRow(
        RedSalamander::DxUi::Label* label, RedSalamander::DxUi::Control* field, float left, float fieldLeft, float right, float top, float height) noexcept
    {
        if (label)
        {
            label->SetBounds(D2D1::RectF(left, top + 6.0f, fieldLeft - 8.0f, top + height));
        }
        if (field)
        {
            field->SetBounds(D2D1::RectF(fieldLeft, top, right, top + height));
        }
    }

    static void LayoutCheckboxPair(RedSalamander::DxUi::Checkbox* leftCheckbox,
                                   RedSalamander::DxUi::Checkbox* rightCheckbox,
                                   float left,
                                   float mid,
                                   float right,
                                   float top,
                                   float height) noexcept
    {
        if (leftCheckbox)
        {
            leftCheckbox->SetBounds(D2D1::RectF(left, top, mid - 6.0f, top + height));
        }
        if (rightCheckbox)
        {
            rightCheckbox->SetBounds(D2D1::RectF(mid + 6.0f, top, right, top + height));
        }
    }

    [[nodiscard]] MakeFileListSettings ReadOptionsFromUi() const
    {
        MakeFileListSettings options{};
        options.sourceMode   = MakeFileListSourceModeFromIndex(_sourceCombo ? _sourceCombo->GetSelectedIndex() : std::nullopt);
        options.recursive    = _recursiveCheckbox && _recursiveCheckbox->IsChecked();
        options.format       = MakeFileListFormatFromIndex(_formatCombo ? _formatCombo->GetSelectedIndex() : std::nullopt);
        options.outputTarget = MakeFileListOutputTargetFromIndex(_outputCombo ? _outputCombo->GetSelectedIndex() : std::nullopt);
        options.textMacro    = _textMacroField ? std::wstring(_textMacroField->GetText()) : std::wstring{};
        options.outputFile =
            std::filesystem::path(StringUtils::TrimWhitespaceCopy(_outputFileField ? std::wstring(_outputFileField->GetText()) : std::wstring{}));
        options.includeName        = _includeNameCheckbox && _includeNameCheckbox->IsChecked();
        options.includeFullPath    = _includeFullPathCheckbox && _includeFullPathCheckbox->IsChecked();
        options.includeSize        = _includeSizeCheckbox && _includeSizeCheckbox->IsChecked();
        options.includeModified    = _includeModifiedCheckbox && _includeModifiedCheckbox->IsChecked();
        options.includeAttributes  = _includeAttributesCheckbox && _includeAttributesCheckbox->IsChecked();
        options.includeDirectories = _includeDirectoriesCheckbox && _includeDirectoriesCheckbox->IsChecked();
        return options;
    }

    void Confirm() noexcept
    {
        _options = ReadOptionsFromUi();
        _result  = _options;
        _done    = true;
        if (_hWnd && IsWindow(_hWnd.get()) != FALSE)
        {
            _hWnd.reset();
        }
    }

    void Cancel() noexcept
    {
        _result.reset();
        _done = true;
        if (_hWnd && IsWindow(_hWnd.get()) != FALSE)
        {
            _hWnd.reset();
        }
    }

    void SetState(const MakeFileListSettings& options) noexcept
    {
        if (_sourceCombo)
        {
            _sourceCombo->SetSelectedIndex(MakeFileListSourceModeToIndex(options.sourceMode));
        }
        if (_recursiveCheckbox)
        {
            _recursiveCheckbox->SetChecked(options.recursive);
        }
        if (_formatCombo)
        {
            _formatCombo->SetSelectedIndex(MakeFileListFormatToIndex(options.format));
        }
        if (_outputCombo)
        {
            _outputCombo->SetSelectedIndex(MakeFileListOutputTargetToIndex(options.outputTarget));
        }
        if (_textMacroField)
        {
            _textMacroField->SetText(options.textMacro);
        }
        if (_outputFileField)
        {
            _outputFileField->SetText(options.outputFile.wstring());
        }
        if (_includeNameCheckbox)
        {
            _includeNameCheckbox->SetChecked(options.includeName);
        }
        if (_includeFullPathCheckbox)
        {
            _includeFullPathCheckbox->SetChecked(options.includeFullPath);
        }
        if (_includeSizeCheckbox)
        {
            _includeSizeCheckbox->SetChecked(options.includeSize);
        }
        if (_includeModifiedCheckbox)
        {
            _includeModifiedCheckbox->SetChecked(options.includeModified);
        }
        if (_includeAttributesCheckbox)
        {
            _includeAttributesCheckbox->SetChecked(options.includeAttributes);
        }
        if (_includeDirectoriesCheckbox)
        {
            _includeDirectoriesCheckbox->SetChecked(options.includeDirectories);
        }
        UpdateOutputEnabledState();
    }

#ifdef ENABLE_TESTS
    LRESULT OnDebugCommand(MakeFileListOptionsPromptDebugCommand command, LPARAM lParam)
    {
        switch (command)
        {
            case MakeFileListOptionsPromptDebugCommand::GetSnapshot:
            {
                auto* snapshot = reinterpret_cast<MakeFileListOptionsPromptDebugSnapshot*>(lParam);
                if (! snapshot)
                {
                    return FALSE;
                }

                const MakeFileListSettings options       = ReadOptionsFromUi();
                *snapshot                                = MakeFileListOptionsPromptDebugSnapshot{};
                snapshot->usesDxUiHost                   = _dxHost.GetRoot() != nullptr;
                snapshot->visibleChildWindowCount        = CountVisibleChildWindowsLocal(_hWnd.get());
                snapshot->visibleNativeChildControlCount = CountVisibleNativeChildControlWindowsLocal(_hWnd.get());
                snapshot->dialogClassName                = GetWindowClassNameLocal(_hWnd.get());
                snapshot->sourceMode                     = static_cast<uint8_t>(options.sourceMode);
                snapshot->recursive                      = options.recursive;
                snapshot->format                         = static_cast<uint8_t>(options.format);
                snapshot->outputTarget                   = static_cast<uint8_t>(options.outputTarget);
                snapshot->textMacro                      = options.textMacro;
                snapshot->outputFileText                 = _outputFileField ? std::wstring(_outputFileField->GetText()) : std::wstring{};
                snapshot->includeName                    = options.includeName;
                snapshot->includeFullPath                = options.includeFullPath;
                snapshot->includeSize                    = options.includeSize;
                snapshot->includeModified                = options.includeModified;
                snapshot->includeAttributes              = options.includeAttributes;
                snapshot->includeDirectories             = options.includeDirectories;
                snapshot->outputFileFieldEnabled         = _outputFileField && _outputFileField->IsEnabled();
                return TRUE;
            }
            case MakeFileListOptionsPromptDebugCommand::SetState:
            {
                const auto* payload = reinterpret_cast<const MakeFileListOptionsPromptDebugStatePayload*>(lParam);
                if (! payload)
                {
                    return FALSE;
                }

                MakeFileListSettings options{};
                options.sourceMode         = MakeFileListSourceModeFromRaw(payload->sourceMode);
                options.recursive          = payload->recursive;
                options.format             = MakeFileListFormatFromRaw(payload->format);
                options.outputTarget       = MakeFileListOutputTargetFromRaw(payload->outputTarget);
                options.textMacro          = payload->textMacro;
                options.outputFile         = std::filesystem::path(payload->outputFile);
                options.includeName        = payload->includeName;
                options.includeFullPath    = payload->includeFullPath;
                options.includeSize        = payload->includeSize;
                options.includeModified    = payload->includeModified;
                options.includeAttributes  = payload->includeAttributes;
                options.includeDirectories = payload->includeDirectories;
                SetState(options);
                return TRUE;
            }
            case MakeFileListOptionsPromptDebugCommand::Confirm: Confirm(); return TRUE;
            case MakeFileListOptionsPromptDebugCommand::Cancel: Cancel(); return TRUE;
        }

        return FALSE;
    }
#endif

private:
    HWND _ownerWindow        = nullptr;
    HWND _restoreFocusWindow = nullptr;
    AppTheme _theme{};
    MakeFileListSettings _options{};
    wil::unique_hwnd _hWnd;
    RedSalamander::DxUi::WindowHost _dxHost;
    RedSalamander::DxUi::ThemePalette _palette{};
    std::unique_ptr<RedSalamander::DxUi::Panel> _rootStorage;
    RedSalamander::DxUi::Panel* _root                          = nullptr;
    RedSalamander::DxUi::Label* _sourceLabel                   = nullptr;
    RedSalamander::DxUi::ComboBox* _sourceCombo                = nullptr;
    RedSalamander::DxUi::Checkbox* _recursiveCheckbox          = nullptr;
    RedSalamander::DxUi::Label* _formatLabel                   = nullptr;
    RedSalamander::DxUi::ComboBox* _formatCombo                = nullptr;
    RedSalamander::DxUi::Label* _textMacroLabel                = nullptr;
    RedSalamander::DxUi::TextField* _textMacroField            = nullptr;
    RedSalamander::DxUi::Label* _fieldsLabel                   = nullptr;
    RedSalamander::DxUi::Checkbox* _includeNameCheckbox        = nullptr;
    RedSalamander::DxUi::Checkbox* _includeFullPathCheckbox    = nullptr;
    RedSalamander::DxUi::Checkbox* _includeSizeCheckbox        = nullptr;
    RedSalamander::DxUi::Checkbox* _includeModifiedCheckbox    = nullptr;
    RedSalamander::DxUi::Checkbox* _includeAttributesCheckbox  = nullptr;
    RedSalamander::DxUi::Checkbox* _includeDirectoriesCheckbox = nullptr;
    RedSalamander::DxUi::Label* _outputLabel                   = nullptr;
    RedSalamander::DxUi::ComboBox* _outputCombo                = nullptr;
    RedSalamander::DxUi::Label* _outputFileLabel               = nullptr;
    RedSalamander::DxUi::TextField* _outputFileField           = nullptr;
    RedSalamander::DxUi::Button* _okButton                     = nullptr;
    RedSalamander::DxUi::Button* _cancelButton                 = nullptr;
    bool _done                                                 = false;
    std::optional<MakeFileListSettings> _result;
};

[[nodiscard]] std::optional<MakeFileListSettings> PromptForMakeFileListOptions(HWND ownerWindow,
                                                                               const AppTheme& theme,
                                                                               const MakeFileListSettings& initialOptions) noexcept
{
    MakeFileListOptionsPromptWindow prompt(ownerWindow, theme, initialOptions);
    return prompt.ShowModal();
}

constexpr size_t kOpenedFilesNoSelection       = static_cast<size_t>(-1);
constexpr size_t kSharedDirectoriesNoSelection = static_cast<size_t>(-1);

[[nodiscard]] std::wstring LoadStringOrFallback(UINT stringId, std::wstring_view fallback)
{
    std::wstring text = LoadStringResource(nullptr, stringId);
    if (text.empty())
    {
        text.assign(fallback);
    }
    return text;
}

constexpr wchar_t kArchivePackPromptClassName[]   = L"RedSalamander.ArchivePackPrompt";
constexpr wchar_t kArchiveUnpackPromptClassName[] = L"RedSalamander.ArchiveUnpackPrompt";

struct ArchivePackerDefinition final
{
    std::wstring displayName;
    std::wstring extensionNoDot;
    GUID classId{};
    bool hasClassId = false;
    bool storedZip  = false;
};

struct ArchivePackPromptResult final
{
    std::filesystem::path archivePath;
    ArchivePackerDefinition packer;
    bool deleteSources = false;
};

struct ArchiveUnpackerDefinition final
{
    std::wstring displayName;
    std::wstring extensionNoDot;
    bool storedZip = false;
};

enum class ArchiveExistingTargetPolicy : uint8_t
{
    Skip,
    Replace,
};

struct ArchiveUnpackPromptResult final
{
    std::filesystem::path destinationPath;
    ArchiveUnpackerDefinition unpacker;
    ArchiveExistingTargetPolicy conflictPolicy = ArchiveExistingTargetPolicy::Skip;
    bool deleteArchive                         = false;
    std::wstring maskText;
};

struct SevenZipArchiveExports final
{
    wil::unique_hmodule module;
    Func_CreateObject createObject               = nullptr;
    Func_GetNumberOfFormats getNumberOfFormats   = nullptr;
    Func_GetHandlerProperty2 getHandlerProperty2 = nullptr;
};

#ifdef ENABLE_TESTS
std::atomic<HWND> g_archivePackPromptWindow{nullptr};
std::atomic<HWND> g_archiveUnpackPromptWindow{nullptr};

enum class ArchivePackPromptDebugCommand : uintptr_t
{
    GetSnapshot = 1u,
    SetPackerIndex,
    SetArchivePath,
    SetDeleteAfter,
    Confirm,
    Cancel,
};

struct ArchivePackPromptArchivePathPayload final
{
    std::wstring archivePath;
};

enum class ArchiveUnpackPromptDebugCommand : uintptr_t
{
    GetSnapshot = 1u,
    SetDestinationPath,
    SetMask,
    SetConflictPolicy,
    SetDeleteAfter,
    Confirm,
    Cancel,
};

struct ArchiveUnpackPromptTextPayload final
{
    std::wstring text;
};

[[nodiscard]] UINT GetArchivePackPromptDebugMessage() noexcept
{
    static const UINT message = RegisterWindowMessageW(L"RedSalamander.ArchivePackPrompt.Debug");
    return message;
}

[[nodiscard]] UINT GetArchiveUnpackPromptDebugMessage() noexcept
{
    static const UINT message = RegisterWindowMessageW(L"RedSalamander.ArchiveUnpackPrompt.Debug");
    return message;
}

#endif

[[nodiscard]] std::wstring ToLowerInvariant(std::wstring text)
{
    std::transform(text.begin(), text.end(), text.begin(), [](wchar_t ch) noexcept { return static_cast<wchar_t>(std::towlower(static_cast<wint_t>(ch))); });
    return text;
}

[[nodiscard]] std::wstring ToUpperInvariant(std::wstring text)
{
    std::transform(text.begin(), text.end(), text.begin(), [](wchar_t ch) noexcept { return static_cast<wchar_t>(std::towupper(static_cast<wint_t>(ch))); });
    return text;
}

[[nodiscard]] std::wstring PropVariantToWideStringForArchive(const PROPVARIANT& value) noexcept
{
    if (value.vt == VT_BSTR && value.bstrVal)
    {
        const UINT length = SysStringLen(value.bstrVal);
        return std::wstring(value.bstrVal, value.bstrVal + length);
    }
    if (value.vt == VT_LPWSTR && value.pwszVal)
    {
        return std::wstring(value.pwszVal);
    }
    return {};
}

[[nodiscard]] bool PropVariantToBoolForArchive(const PROPVARIANT& value) noexcept
{
    return value.vt == VT_BOOL && value.boolVal != VARIANT_FALSE;
}

[[nodiscard]] bool PropVariantToUInt64ForArchive(const PROPVARIANT& value, uint64_t& outValue) noexcept
{
    outValue = 0u;
    if (value.vt == VT_UI8)
    {
        outValue = static_cast<uint64_t>(value.uhVal.QuadPart);
        return true;
    }
    if (value.vt == VT_UI4)
    {
        outValue = static_cast<uint64_t>(value.ulVal);
        return true;
    }
    if (value.vt == VT_I8 && value.hVal.QuadPart >= 0)
    {
        outValue = static_cast<uint64_t>(value.hVal.QuadPart);
        return true;
    }
    if (value.vt == VT_I4 && value.lVal >= 0)
    {
        outValue = static_cast<uint64_t>(value.lVal);
        return true;
    }
    return false;
}

[[nodiscard]] std::optional<GUID> PropVariantToGuidBinaryBstrForArchive(const PROPVARIANT& value) noexcept
{
    if (value.vt != VT_BSTR || ! value.bstrVal)
    {
        return std::nullopt;
    }

    const UINT bytes = SysStringByteLen(value.bstrVal);
    if (bytes != sizeof(GUID))
    {
        return std::nullopt;
    }

    GUID guid{};
    std::memcpy(&guid, value.bstrVal, sizeof(guid));
    return guid;
}

[[nodiscard]] std::wstring FirstArchiveExtensionToken(std::wstring_view extensionList)
{
    size_t pos = 0u;
    while (pos < extensionList.size())
    {
        while (pos < extensionList.size() && std::iswspace(static_cast<wint_t>(extensionList[pos])) != 0)
        {
            ++pos;
        }

        const size_t start = pos;
        while (pos < extensionList.size() && std::iswspace(static_cast<wint_t>(extensionList[pos])) == 0)
        {
            ++pos;
        }

        if (start == pos)
        {
            break;
        }

        std::wstring token(extensionList.substr(start, pos - start));
        while (! token.empty() && token.front() == L'.')
        {
            token.erase(token.begin());
        }
        if (! token.empty())
        {
            return ToLowerInvariant(std::move(token));
        }
    }

    return {};
}

[[nodiscard]] bool ArchiveExtensionListContains(std::wstring_view extensionList, std::wstring_view extensionNoDotLower) noexcept
{
    if (extensionList.empty() || extensionNoDotLower.empty())
    {
        return false;
    }

    size_t pos = 0u;
    while (pos < extensionList.size())
    {
        while (pos < extensionList.size() && std::iswspace(static_cast<wint_t>(extensionList[pos])) != 0)
        {
            ++pos;
        }

        const size_t start = pos;
        while (pos < extensionList.size() && std::iswspace(static_cast<wint_t>(extensionList[pos])) == 0)
        {
            ++pos;
        }

        if (start == pos)
        {
            break;
        }

        std::wstring_view token(extensionList.data() + start, pos - start);
        while (! token.empty() && token.front() == L'.')
        {
            token.remove_prefix(1);
        }

        if (token.size() == extensionNoDotLower.size() && OrdinalString::EqualsNoCase(token, extensionNoDotLower))
        {
            return true;
        }
    }

    return false;
}

[[nodiscard]] std::wstring ArchiveExtensionNoDotLower(const std::filesystem::path& archivePath)
{
    std::wstring extension = archivePath.extension().wstring();
    if (! extension.empty() && extension.front() == L'.')
    {
        extension.erase(extension.begin());
    }
    return ToLowerInvariant(std::move(extension));
}

[[nodiscard]] std::filesystem::path GetExecutableDirectoryForArchivePackers() noexcept
{
    std::array<wchar_t, MAX_PATH> stackBuffer{};
    DWORD chars = GetModuleFileNameW(nullptr, stackBuffer.data(), static_cast<DWORD>(stackBuffer.size()));
    if (chars == 0u)
    {
        return {};
    }

    if (chars < stackBuffer.size())
    {
        return std::filesystem::path(stackBuffer.data()).parent_path();
    }

    std::vector<wchar_t> buffer(stackBuffer.size() * 2u, L'\0');
    for (;;)
    {
        chars = GetModuleFileNameW(nullptr, buffer.data(), static_cast<DWORD>(buffer.size()));
        if (chars == 0u)
        {
            return {};
        }
        if (chars < buffer.size())
        {
            return std::filesystem::path(buffer.data()).parent_path();
        }
        buffer.assign(buffer.size() * 2u, L'\0');
    }
}

[[nodiscard]] HRESULT LoadSevenZipArchiveExports(SevenZipArchiveExports& out) noexcept
{
    if (out.module && out.createObject && out.getNumberOfFormats && out.getHandlerProperty2)
    {
        return S_OK;
    }

    out.module.reset();
    out.createObject        = nullptr;
    out.getNumberOfFormats  = nullptr;
    out.getHandlerProperty2 = nullptr;

    const std::filesystem::path exeDir = GetExecutableDirectoryForArchivePackers();
    if (exeDir.empty())
    {
        return HRESULT_FROM_WIN32(ERROR_NOT_FOUND);
    }

    const std::array<std::filesystem::path, 2> candidates{exeDir / L"Plugins" / L"7zip.dll", exeDir / L"7zip.dll"};
    wil::unique_hmodule module;
    for (const std::filesystem::path& candidate : candidates)
    {
        module.reset(LoadLibraryW(candidate.c_str()));
        if (module)
        {
            break;
        }
    }
    if (! module)
    {
        const DWORD lastError = GetLastError();
        return HRESULT_FROM_WIN32(lastError != ERROR_SUCCESS ? lastError : ERROR_MOD_NOT_FOUND);
    }

#pragma warning(push)
#pragma warning(disable : 4191) // 7-Zip exports are documented C entry points with these signatures.
    const auto createObject        = reinterpret_cast<Func_CreateObject>(GetProcAddress(module.get(), "CreateObject"));
    const auto getNumberOfFormats  = reinterpret_cast<Func_GetNumberOfFormats>(GetProcAddress(module.get(), "GetNumberOfFormats"));
    const auto getHandlerProperty2 = reinterpret_cast<Func_GetHandlerProperty2>(GetProcAddress(module.get(), "GetHandlerProperty2"));
#pragma warning(pop)

    if (! createObject || ! getNumberOfFormats || ! getHandlerProperty2)
    {
        return HRESULT_FROM_WIN32(ERROR_PROC_NOT_FOUND);
    }

    out.module              = std::move(module);
    out.createObject        = createObject;
    out.getNumberOfFormats  = getNumberOfFormats;
    out.getHandlerProperty2 = getHandlerProperty2;
    return S_OK;
}

[[nodiscard]] SevenZipArchiveExports& GetSevenZipArchiveExports() noexcept
{
    static SevenZipArchiveExports exports;
    return exports;
}

[[nodiscard]] std::optional<GUID> TryGetSevenZipArchiveClassIdForExtension(SevenZipArchiveExports& api, std::wstring_view extensionNoDotLower) noexcept
{
    if (extensionNoDotLower.empty() || ! api.getNumberOfFormats || ! api.getHandlerProperty2)
    {
        return std::nullopt;
    }

    UInt32 formatCount = 0u;
    if (FAILED(api.getNumberOfFormats(&formatCount)))
    {
        return std::nullopt;
    }

    for (UInt32 index = 0u; index < formatCount; ++index)
    {
        PROPVARIANT extensionVar{};
        PropVariantInit(&extensionVar);
        const HRESULT extensionHr        = api.getHandlerProperty2(index, NArchive::NHandlerPropID::kExtension, &extensionVar);
        auto clearExtension              = wil::scope_exit([&]
        {
            if (SUCCEEDED(extensionHr))
            {
                PropVariantClear(&extensionVar);
            }
        });
        const std::wstring extensionList = SUCCEEDED(extensionHr) ? PropVariantToWideStringForArchive(extensionVar) : std::wstring{};

        PROPVARIANT addExtensionVar{};
        PropVariantInit(&addExtensionVar);
        const HRESULT addExtensionHr        = api.getHandlerProperty2(index, NArchive::NHandlerPropID::kAddExtension, &addExtensionVar);
        auto clearAddExtension              = wil::scope_exit([&]
        {
            if (SUCCEEDED(addExtensionHr))
            {
                PropVariantClear(&addExtensionVar);
            }
        });
        const std::wstring addExtensionList = SUCCEEDED(addExtensionHr) ? PropVariantToWideStringForArchive(addExtensionVar) : std::wstring{};

        if (! ArchiveExtensionListContains(extensionList, extensionNoDotLower) && ! ArchiveExtensionListContains(addExtensionList, extensionNoDotLower))
        {
            continue;
        }

        PROPVARIANT classIdVar{};
        PropVariantInit(&classIdVar);
        if (FAILED(api.getHandlerProperty2(index, NArchive::NHandlerPropID::kClassID, &classIdVar)))
        {
            continue;
        }
        auto clearClassId                 = wil::scope_exit([&] { PropVariantClear(&classIdVar); });
        const std::optional<GUID> classId = PropVariantToGuidBinaryBstrForArchive(classIdVar);
        if (classId.has_value())
        {
            return classId;
        }
    }

    return std::nullopt;
}

[[nodiscard]] ArchivePackerDefinition BuildStoredZipPackerDefinition()
{
    return ArchivePackerDefinition{.displayName = L"ZIP (Plugin)", .extensionNoDot = L"zip", .storedZip = true};
}

[[nodiscard]] ArchiveUnpackerDefinition BuildStoredZipUnpackerDefinition()
{
    return ArchiveUnpackerDefinition{.displayName = L"ZIP (Plugin)", .extensionNoDot = L"zip", .storedZip = true};
}

[[nodiscard]] std::vector<ArchivePackerDefinition> EnumerateArchivePackerDefinitions()
{
    std::vector<ArchivePackerDefinition> packers;
    packers.push_back(BuildStoredZipPackerDefinition());

    SevenZipArchiveExports& api = GetSevenZipArchiveExports();
    if (FAILED(LoadSevenZipArchiveExports(api)) || ! api.getNumberOfFormats || ! api.getHandlerProperty2)
    {
        return packers;
    }

    UInt32 formatCount = 0u;
    if (FAILED(api.getNumberOfFormats(&formatCount)))
    {
        return packers;
    }

    std::unordered_set<std::wstring> seenExtensions;
    seenExtensions.insert(L"zip");

    for (UInt32 index = 0u; index < formatCount; ++index)
    {
        PROPVARIANT updateVar{};
        PropVariantInit(&updateVar);
        if (FAILED(api.getHandlerProperty2(index, NArchive::NHandlerPropID::kUpdate, &updateVar)))
        {
            continue;
        }
        auto clearUpdate = wil::scope_exit([&] { PropVariantClear(&updateVar); });
        if (! PropVariantToBoolForArchive(updateVar))
        {
            continue;
        }

        PROPVARIANT extensionVar{};
        PropVariantInit(&extensionVar);
        if (FAILED(api.getHandlerProperty2(index, NArchive::NHandlerPropID::kExtension, &extensionVar)))
        {
            continue;
        }
        auto clearExtension    = wil::scope_exit([&] { PropVariantClear(&extensionVar); });
        std::wstring extension = FirstArchiveExtensionToken(PropVariantToWideStringForArchive(extensionVar));
        if (extension.empty())
        {
            PROPVARIANT addExtensionVar{};
            PropVariantInit(&addExtensionVar);
            const HRESULT addExtensionHr = api.getHandlerProperty2(index, NArchive::NHandlerPropID::kAddExtension, &addExtensionVar);
            auto clearAddExtension       = wil::scope_exit([&]
            {
                if (SUCCEEDED(addExtensionHr))
                {
                    PropVariantClear(&addExtensionVar);
                }
            });
            if (SUCCEEDED(addExtensionHr))
            {
                extension = FirstArchiveExtensionToken(PropVariantToWideStringForArchive(addExtensionVar));
            }
        }
        if (extension.empty() || seenExtensions.contains(extension))
        {
            continue;
        }

        PROPVARIANT classIdVar{};
        PropVariantInit(&classIdVar);
        if (FAILED(api.getHandlerProperty2(index, NArchive::NHandlerPropID::kClassID, &classIdVar)))
        {
            continue;
        }
        auto clearClassId                 = wil::scope_exit([&] { PropVariantClear(&classIdVar); });
        const std::optional<GUID> classId = PropVariantToGuidBinaryBstrForArchive(classIdVar);
        if (! classId.has_value())
        {
            continue;
        }

        PROPVARIANT nameVar{};
        PropVariantInit(&nameVar);
        const HRESULT nameHr = api.getHandlerProperty2(index, NArchive::NHandlerPropID::kName, &nameVar);
        auto clearName       = wil::scope_exit([&]
        {
            if (SUCCEEDED(nameHr))
            {
                PropVariantClear(&nameVar);
            }
        });

        std::wstring name = SUCCEEDED(nameHr) ? PropVariantToWideStringForArchive(nameVar) : std::wstring{};
        if (name.empty())
        {
            name = ToUpperInvariant(extension);
        }

        ArchivePackerDefinition packer{};
        packer.displayName    = std::format(L"{} (7-Zip)", name);
        packer.extensionNoDot = extension;
        packer.classId        = classId.value();
        packer.hasClassId     = true;
        packers.push_back(std::move(packer));
        seenExtensions.insert(std::move(extension));
    }

    return packers;
}

[[nodiscard]] std::vector<RedSalamander::DxUi::ComboBox::Item> BuildArchivePackerComboItems(const std::vector<ArchivePackerDefinition>& packers)
{
    std::vector<RedSalamander::DxUi::ComboBox::Item> items;
    items.reserve(packers.size());
    for (const ArchivePackerDefinition& packer : packers)
    {
        items.push_back(RedSalamander::DxUi::ComboBox::Item{packer.extensionNoDot, packer.displayName});
    }
    return items;
}

[[nodiscard]] std::vector<RedSalamander::DxUi::ComboBox::Item> BuildArchiveUnpackerComboItems(const std::vector<ArchiveUnpackerDefinition>& unpackers)
{
    std::vector<RedSalamander::DxUi::ComboBox::Item> items;
    items.reserve(unpackers.size());
    for (const ArchiveUnpackerDefinition& unpacker : unpackers)
    {
        items.push_back(RedSalamander::DxUi::ComboBox::Item{unpacker.extensionNoDot, unpacker.displayName});
    }
    return items;
}

[[nodiscard]] std::wstring SanitizeArchiveSuggestedStem(std::wstring text)
{
    for (wchar_t& ch : text)
    {
        if (ch < 32 || ch == L'<' || ch == L'>' || ch == L':' || ch == L'"' || ch == L'/' || ch == L'\\' || ch == L'|' || ch == L'?' || ch == L'*')
        {
            ch = L'_';
        }
    }

    text = StringUtils::TrimWhitespaceCopy(text);
    while (! text.empty() && (text.back() == L'.' || text.back() == L' '))
    {
        text.pop_back();
    }
    return text.empty() ? std::wstring(L"selection") : text;
}

[[nodiscard]] std::wstring BuildArchiveSuggestedStem(const std::filesystem::path& currentFolder, const std::vector<std::filesystem::path>& selectedPaths)
{
    std::wstring stem;
    if (selectedPaths.size() == 1u)
    {
        std::error_code ec;
        const std::filesystem::file_status status = std::filesystem::symlink_status(selectedPaths.front(), ec);
        if (! ec && std::filesystem::is_regular_file(status))
        {
            stem = selectedPaths.front().stem().wstring();
        }
        else
        {
            stem = selectedPaths.front().filename().wstring();
        }
    }
    else
    {
        stem = currentFolder.filename().wstring();
    }

    return SanitizeArchiveSuggestedStem(std::move(stem));
}

[[nodiscard]] std::filesystem::path ComposeArchiveCandidatePath(const std::filesystem::path& folder,
                                                                std::wstring_view stem,
                                                                std::wstring_view extensionNoDot,
                                                                unsigned int suffix)
{
    std::wstring filename(stem);
    if (suffix != 0u)
    {
        filename.append(std::format(L" ({})", suffix));
    }

    std::wstring extension(extensionNoDot);
    while (! extension.empty() && extension.front() == L'.')
    {
        extension.erase(extension.begin());
    }
    if (! extension.empty())
    {
        filename.push_back(L'.');
        filename.append(extension);
    }

    return folder / filename;
}

[[nodiscard]] std::filesystem::path ResolveUniqueArchivePath(const std::filesystem::path& folder, std::wstring_view stem, std::wstring_view extensionNoDot)
{
    const auto isTaken = [](const std::filesystem::path& candidate) noexcept
    {
        std::error_code ec;
        const bool exists = std::filesystem::exists(candidate, ec);
        return exists || static_cast<bool>(ec);
    };

    std::filesystem::path candidate = ComposeArchiveCandidatePath(folder, stem, extensionNoDot, 0u);
    if (! isTaken(candidate))
    {
        return candidate;
    }

    for (unsigned int suffix = 2u; suffix < 10000u; ++suffix)
    {
        candidate = ComposeArchiveCandidatePath(folder, stem, extensionNoDot, suffix);
        if (! isTaken(candidate))
        {
            return candidate;
        }
    }

    return ComposeArchiveCandidatePath(folder, std::format(L"{}-{}", stem, GetTickCount64()), extensionNoDot, 0u);
}

[[nodiscard]] std::filesystem::path ComposeArchiveDestinationCandidatePath(const std::filesystem::path& folder, std::wstring_view stem, unsigned int suffix)
{
    std::wstring name(stem);
    if (suffix != 0u)
    {
        name.append(std::format(L" ({})", suffix));
    }
    return folder / name;
}

[[nodiscard]] std::filesystem::path ResolveUniqueArchiveDestinationPath(const std::filesystem::path& folder, std::wstring_view stem)
{
    const auto isTaken = [](const std::filesystem::path& candidate) noexcept
    {
        std::error_code ec;
        const bool exists = std::filesystem::exists(candidate, ec);
        return exists || static_cast<bool>(ec);
    };

    std::filesystem::path candidate = ComposeArchiveDestinationCandidatePath(folder, stem, 0u);
    if (! isTaken(candidate))
    {
        return candidate;
    }

    for (unsigned int suffix = 2u; suffix < 10000u; ++suffix)
    {
        candidate = ComposeArchiveDestinationCandidatePath(folder, stem, suffix);
        if (! isTaken(candidate))
        {
            return candidate;
        }
    }

    return ComposeArchiveDestinationCandidatePath(folder, std::format(L"{}-{}", stem, GetTickCount64()), 0u);
}

[[nodiscard]] std::filesystem::path BuildArchiveUnpackSuggestedDestinationPath(const std::filesystem::path& archivePath,
                                                                               const std::filesystem::path& currentFolder)
{
    std::filesystem::path folder = archivePath.parent_path();
    if (folder.empty())
    {
        folder = currentFolder;
    }
    if (folder.empty())
    {
        std::error_code ec;
        folder = std::filesystem::current_path(ec);
        if (ec)
        {
            return {};
        }
    }

    std::wstring stem = SanitizeArchiveSuggestedStem(archivePath.stem().wstring());
    if (stem.empty())
    {
        stem = L"archive";
    }
    return ResolveUniqueArchiveDestinationPath(folder, stem);
}

[[nodiscard]] std::filesystem::path EnsureArchivePathExtension(const std::filesystem::path& path, std::wstring_view extensionNoDot)
{
    std::filesystem::path result = path;
    std::wstring extension(extensionNoDot);
    while (! extension.empty() && extension.front() == L'.')
    {
        extension.erase(extension.begin());
    }
    if (! extension.empty())
    {
        result.replace_extension(L"." + extension);
    }
    return result;
}

[[nodiscard]] std::vector<RedSalamander::DxUi::ComboBox::Item> BuildArchiveConflictPolicyComboItems()
{
    return {
        {LoadStringResource(nullptr, IDS_ARCHIVE_UNPACK_CONFLICT_SKIP), L"skip"},
        {LoadStringResource(nullptr, IDS_ARCHIVE_UNPACK_CONFLICT_REPLACE), L"replace"},
    };
}

[[nodiscard]] HRESULT ValidateArchiveOutputOutsideSelectedSources(const std::vector<std::filesystem::path>& selectedPaths,
                                                                  const std::filesystem::path& archivePath) noexcept
{
    std::error_code ec;
    const std::filesystem::path normalizedArchive = std::filesystem::absolute(archivePath, ec).lexically_normal();
    if (ec || normalizedArchive.empty())
    {
        return ec ? HRESULT_FROM_WIN32(static_cast<DWORD>(ec.value())) : HRESULT_FROM_WIN32(ERROR_INVALID_NAME);
    }

    for (const std::filesystem::path& selectedPath : selectedPaths)
    {
        ec.clear();
        const std::filesystem::path normalizedSource = std::filesystem::absolute(selectedPath, ec).lexically_normal();
        if (ec || normalizedSource.empty())
        {
            return ec ? HRESULT_FROM_WIN32(static_cast<DWORD>(ec.value())) : HRESULT_FROM_WIN32(ERROR_INVALID_NAME);
        }

        const DWORD attributes = GetFileAttributesW(selectedPath.c_str());
        if (attributes == INVALID_FILE_ATTRIBUTES)
        {
            return HRESULT_FROM_WIN32(GetLastError());
        }

        const std::wstring_view sourceText  = normalizedSource.native();
        const std::wstring_view archiveText = normalizedArchive.native();
        const bool samePath                 = Common::Paths::NormalizedWindowsPathEqualsNoCase(sourceText, archiveText);
        const bool sourceIsDirectory        = (attributes & FILE_ATTRIBUTE_DIRECTORY) != 0u;
        if (samePath || (sourceIsDirectory && Common::Paths::IsSameOrDescendantNormalizedWindowsPath(sourceText, archiveText)))
        {
            return HRESULT_FROM_WIN32(ERROR_ACCESS_DENIED);
        }
    }

    return S_OK;
}

[[nodiscard]] std::wstring ShortDisplayNameForArchiveSelection(const std::vector<std::filesystem::path>& selectedPaths)
{
    if (selectedPaths.size() != 1u)
    {
        return {};
    }

    std::wstring display = selectedPaths.front().filename().wstring();
    if (display.empty())
    {
        display = selectedPaths.front().wstring();
    }
    constexpr size_t kMaxDisplayChars = 46u;
    if (display.size() > kMaxDisplayChars)
    {
        display.resize(kMaxDisplayChars - 1u);
        display.push_back(L'\x2026');
    }
    return display;
}

class ArchivePackPromptWindow final
{
public:
    ArchivePackPromptWindow(const ArchivePackPromptWindow&)            = delete;
    ArchivePackPromptWindow& operator=(const ArchivePackPromptWindow&) = delete;
    ArchivePackPromptWindow(ArchivePackPromptWindow&&)                 = delete;
    ArchivePackPromptWindow& operator=(ArchivePackPromptWindow&&)      = delete;

    ArchivePackPromptWindow(HWND ownerWindow,
                            const AppTheme& theme,
                            std::filesystem::path currentFolder,
                            std::vector<std::filesystem::path> selectedPaths,
                            std::vector<ArchivePackerDefinition> packers) noexcept
        : _ownerWindow(GetOwnerWindowOrSelf(ownerWindow)),
          _restoreFocusWindow(ownerWindow && IsWindow(ownerWindow) != FALSE ? ownerWindow : nullptr),
          _theme(theme),
          _currentFolder(std::move(currentFolder)),
          _selectedPaths(std::move(selectedPaths)),
          _packers(std::move(packers))
    {
        if (_packers.empty())
        {
            _packers.push_back(BuildStoredZipPackerDefinition());
        }

        _suggestedStem = BuildArchiveSuggestedStem(_currentFolder, _selectedPaths);
        if (_ownerWindow && IsWindow(_ownerWindow) != FALSE)
        {
            const HWND focused = GetFocus();
            if (focused && IsWindow(focused) != FALSE && (focused == _ownerWindow || IsChild(_ownerWindow, focused) != FALSE))
            {
                _restoreFocusWindow = focused;
            }
            else if (! _restoreFocusWindow || IsWindow(_restoreFocusWindow) == FALSE ||
                     (_restoreFocusWindow != _ownerWindow && IsChild(_ownerWindow, _restoreFocusWindow) == FALSE))
            {
                _restoreFocusWindow = _ownerWindow;
            }
        }
        else if (! _restoreFocusWindow || IsWindow(_restoreFocusWindow) == FALSE)
        {
            _restoreFocusWindow = nullptr;
        }
    }

    [[nodiscard]] std::optional<ArchivePackPromptResult> ShowModal() noexcept
    {
        const HRESULT classHr = EnsureWindowClass();
        if (FAILED(classHr))
        {
            return std::nullopt;
        }

        const DWORD style   = WS_CAPTION | WS_SYSMENU | WS_POPUP | WS_CLIPCHILDREN | WS_CLIPSIBLINGS;
        const DWORD exStyle = WS_EX_DLGMODALFRAME;
        const UINT dpi      = _ownerWindow && IsWindow(_ownerWindow) != FALSE ? GetDpiForWindow(_ownerWindow) : GetDpiForSystem();

        RECT bounds{0, 0, ScalePanePromptForDpi(dpi, 520), ScalePanePromptForDpi(dpi, 292)};
        if (AdjustWindowRectExForDpi(&bounds, style, FALSE, exStyle, dpi) == FALSE)
        {
            return std::nullopt;
        }

        const bool restoreOwnerEnabled = _ownerWindow && IsWindow(_ownerWindow) != FALSE && IsWindowEnabled(_ownerWindow) != FALSE;
        if (restoreOwnerEnabled)
        {
            EnableWindow(_ownerWindow, FALSE);
        }
        const auto restoreOwner = wil::scope_exit([this, restoreOwnerEnabled]() noexcept
        {
            if (restoreOwnerEnabled && _ownerWindow && IsWindow(_ownerWindow) != FALSE)
            {
                EnableWindow(_ownerWindow, TRUE);
                SetActiveWindow(_ownerWindow);
                const HWND restoreFocus = (_restoreFocusWindow && IsWindow(_restoreFocusWindow) != FALSE &&
                                           (_restoreFocusWindow == _ownerWindow || IsChild(_ownerWindow, _restoreFocusWindow) != FALSE))
                                              ? _restoreFocusWindow
                                              : _ownerWindow;
                SetFocus(restoreFocus);
            }
        });

        const std::wstring caption = LoadStringResource(nullptr, IDS_CMD_PACK);
        const HWND hwnd            = CreateWindowExW(exStyle,
                                                     kArchivePackPromptClassName,
                                                     caption.c_str(),
                                                     style,
                                                     CW_USEDEFAULT,
                                                     CW_USEDEFAULT,
                                                     bounds.right - bounds.left,
                                                     bounds.bottom - bounds.top,
                                                     _ownerWindow,
                                                     nullptr,
                                                     GetModuleHandleW(nullptr),
                                                     this);
        if (! hwnd)
        {
            return std::nullopt;
        }
        if (! _hWnd)
        {
            _hWnd.reset(hwnd);
        }

        static_cast<void>(Common::WindowSizing::CenterExistingWindowOnOwner(_hWnd.get(), _ownerWindow));
        static_cast<void>(_dxHost.PrimeForShow());
        ShowWindow(_hWnd.get(), SW_SHOWNORMAL);
        UpdateWindow(_hWnd.get());
        SetForegroundWindow(_hWnd.get());

        const RedSalamander::DxUi::DxUiModalLoopResult loopResult = RedSalamander::DxUi::RunDxUiModalLoop(_hWnd.get(),
                                                                                                          RedSalamander::DxUi::DxUiModalLoopOptions{
                                                                                                              .diagnosticName = L"ArchivePackPrompt",
                                                                                                              .shouldContinue = ShouldContinueModalLoop,
                                                                                                              .context        = this,
                                                                                                              .onQuit         = OnModalLoopQuit,
                                                                                                          });
        if (loopResult == RedSalamander::DxUi::DxUiModalLoopResult::GetMessageFailed)
        {
            return std::nullopt;
        }

        return _result;
    }

    static LRESULT CALLBACK WndProc(HWND hwnd, UINT message, WPARAM wParam, LPARAM lParam) noexcept
    {
        if (message == WM_NCCREATE)
        {
            auto* cs   = reinterpret_cast<CREATESTRUCTW*>(lParam);
            auto* self = static_cast<ArchivePackPromptWindow*>(cs ? cs->lpCreateParams : nullptr);
            if (! self)
            {
                return FALSE;
            }

            SetWindowLongPtrW(hwnd, GWLP_USERDATA, reinterpret_cast<LONG_PTR>(self));
            if (! self->_hWnd)
            {
                self->_hWnd.reset(hwnd);
            }
            return TRUE;
        }

        auto* self = reinterpret_cast<ArchivePackPromptWindow*>(GetWindowLongPtrW(hwnd, GWLP_USERDATA));
        if (! self)
        {
            return DefWindowProcW(hwnd, message, wParam, lParam);
        }

        bool handled     = false;
        LRESULT dxResult = 0;
        if (message != WM_CREATE)
        {
            dxResult = self->_dxHost.HandleMessage(hwnd, message, wParam, lParam, handled);
        }
        if (handled)
        {
            if (message == WM_SIZE || message == WM_DPICHANGED)
            {
                self->Layout();
            }
            if (message == WM_NCDESTROY)
            {
#ifdef ENABLE_TESTS
                g_archivePackPromptWindow.store(nullptr);
#endif
                SetWindowLongPtrW(hwnd, GWLP_USERDATA, 0);
                self->_dxHost.Detach();
                if (self->_hWnd.get() == hwnd)
                {
                    static_cast<void>(self->_hWnd.release());
                }
                self->_done = true;
            }
            return dxResult;
        }

#ifdef ENABLE_TESTS
        if (message == GetArchivePackPromptDebugMessage())
        {
            return self->OnDebugCommand(static_cast<ArchivePackPromptDebugCommand>(wParam), lParam);
        }
#endif

        switch (message)
        {
            case WM_CREATE: return self->OnCreate(hwnd) ? 0 : -1;
            case WM_SIZE: self->Layout(); return 0;
            case WM_DPICHANGED:
            {
                const auto* suggested = reinterpret_cast<const RECT*>(lParam);
                if (suggested)
                {
                    SetWindowPos(hwnd,
                                 nullptr,
                                 suggested->left,
                                 suggested->top,
                                 suggested->right - suggested->left,
                                 suggested->bottom - suggested->top,
                                 SWP_NOZORDER | SWP_NOACTIVATE);
                }
                self->Layout();
                return 0;
            }
            case WM_ERASEBKGND: return 1;
            case WM_NCACTIVATE: ApplyTitleBarTheme(hwnd, self->_theme, wParam != FALSE); return DefWindowProcW(hwnd, message, wParam, lParam);
            case WM_CLOSE: self->Cancel(); return 0;
            case WM_NCDESTROY:
#ifdef ENABLE_TESTS
                g_archivePackPromptWindow.store(nullptr);
#endif
                self->_dxHost.Detach();
                if (self->_hWnd.get() == hwnd)
                {
                    self->_hWnd.release();
                }
                SetWindowLongPtrW(hwnd, GWLP_USERDATA, 0);
                self->_done = true;
                break;
        }

        return DefWindowProcW(hwnd, message, wParam, lParam);
    }

private:
    [[nodiscard]] static HRESULT EnsureWindowClass() noexcept
    {
        static ATOM atom = 0;
        if (atom != 0)
        {
            return S_OK;
        }

        WNDCLASSEXW wc{};
        wc.cbSize        = sizeof(wc);
        wc.lpfnWndProc   = ArchivePackPromptWindow::WndProc;
        wc.hInstance     = GetModuleHandleW(nullptr);
        wc.hCursor       = LoadCursorW(nullptr, IDC_ARROW);
        wc.lpszClassName = kArchivePackPromptClassName;
        wc.style         = CS_DBLCLKS;

        atom = RegisterClassExW(&wc);
        return atom != 0 ? S_OK : HRESULT_FROM_WIN32(GetLastError());
    }

    [[nodiscard]] bool OnCreate(HWND hwnd) noexcept
    {
        if (! _dxHost.Attach(hwnd))
        {
            return false;
        }

#ifdef ENABLE_TESTS
        g_archivePackPromptWindow.store(hwnd);
#endif
        BuildUi();
        ApplyTheme();
        ResetArchivePathForSelectedPacker();
        Layout();
        if (_archivePathCombo)
        {
            _dxHost.SetFocusControl(_archivePathCombo);
        }
        _dxHost.SetDefaultButton(_okButton);
        _dxHost.SetCancelButton(_cancelButton);
        return true;
    }

    void BuildUi()
    {
        if (_root)
        {
            return;
        }

        using namespace RedSalamander::DxUi;

        _rootStorage = std::make_unique<Panel>();
        _root        = _rootStorage.get();

        const std::wstring selectedName = ShortDisplayNameForArchiveSelection(_selectedPaths);
        const UINT promptStringId = _selectedPaths.size() == 1u ? (IsSingleSelectedDirectory() ? IDS_ARCHIVE_PACK_DIRECTORY_FMT : IDS_ARCHIVE_PACK_FILE_FMT)
                                                                : IDS_ARCHIVE_PACK_SELECTION_FMT;
        std::wstring promptText =
            _selectedPaths.size() == 1u ? FormatStringResource(nullptr, promptStringId, selectedName) : LoadStringResource(nullptr, promptStringId);
        if (promptText.empty())
        {
            promptText = LoadStringResource(nullptr, IDS_CMD_PACK);
        }

        _promptLabel = _root->AddChild<Label>(std::move(promptText));
        _promptLabel->SetAlignment(DWRITE_TEXT_ALIGNMENT_LEADING);

        _archivePathCombo = _root->AddChild<ComboBox>();
        _archivePathCombo->SetVariant(ComboBoxVariant::Window);
        _archivePathCombo->SetEditable(true);
        _archivePathCombo->SetOnTextChanged([this](std::wstring_view) noexcept
        {
            if (! _updatingArchivePathText)
            {
                _archivePathUserEdited = true;
            }
        });
        _archivePathCombo->SetOnSubmitted([this] { Confirm(); });

        _packerLabel = _root->AddChild<Label>(LoadStringResource(nullptr, IDS_ARCHIVE_PACKER_LABEL));
        _packerLabel->SetAlignment(DWRITE_TEXT_ALIGNMENT_LEADING);

        _packerCombo = _root->AddChild<ComboBox>();
        _packerCombo->SetVariant(ComboBoxVariant::Window);
        _packerCombo->SetEditable(false);
        _packerCombo->SetItems(BuildArchivePackerComboItems(_packers));
        _packerCombo->SetSelectedIndex(0u);
        _packerCombo->SetOnSelectionChanged([this](size_t) noexcept { ResetArchivePathForSelectedPacker(); });
        _packerCombo->SetOnSubmitted([this] { Confirm(); });

        _deleteAfterCheckbox = _root->AddChild<Checkbox>(LoadStringResource(nullptr, IDS_ARCHIVE_PACK_DELETE_AFTER));
        _deleteAfterCheckbox->SetChecked(false);

        _okButton = _root->AddChild<Button>(LoadStringResource(nullptr, IDS_BTN_OK));
        _okButton->SetPrimary(true);
        _okButton->SetOnClick([this] { Confirm(); });

        _cancelButton = _root->AddChild<Button>(LoadStringResource(nullptr, IDS_BTN_CANCEL));
        _cancelButton->SetOnClick([this] { Cancel(); });

        _helpButton = _root->AddChild<Button>(LoadStringResource(nullptr, IDS_BTN_HELP));
        _helpButton->SetOnClick([]() noexcept { MessageBeep(MB_ICONINFORMATION); });

        _dxHost.SetRoot(std::move(_rootStorage));
    }

    [[nodiscard]] bool IsSingleSelectedDirectory() const noexcept
    {
        if (_selectedPaths.size() != 1u)
        {
            return false;
        }

        std::error_code ec;
        const std::filesystem::file_status status = std::filesystem::symlink_status(_selectedPaths.front(), ec);
        return ! ec && std::filesystem::is_directory(status);
    }

    void ApplyTheme() noexcept
    {
        _palette = MakeAppThemeDxPalette(_theme, _theme.windowBackground);
        _dxHost.SetTheme(_palette);
        if (_hWnd)
        {
            ApplyWindowChromeTheme(_hWnd.get(), _theme, WindowBackdropTarget::Tool, GetActiveWindow() == _hWnd.get());
            static_cast<void>(RedrawWindow(_hWnd.get(), nullptr, nullptr, RDW_FRAME | RDW_INVALIDATE | RDW_UPDATENOW));
        }
    }

    [[nodiscard]] size_t SelectedPackerIndex() const noexcept
    {
        const std::optional<size_t> selectedIndex = _packerCombo ? _packerCombo->GetSelectedIndex() : std::optional<size_t>{0u};
        const size_t index                        = selectedIndex.value_or(0u);
        return index < _packers.size() ? index : 0u;
    }

    [[nodiscard]] const ArchivePackerDefinition& SelectedPacker() const noexcept
    {
        return _packers[SelectedPackerIndex()];
    }

    void SetArchivePathText(std::filesystem::path path, bool markUserEdited) noexcept
    {
        if (! _archivePathCombo)
        {
            return;
        }

        const auto restoreUpdateFlag = wil::scope_exit([&]() noexcept { _updatingArchivePathText = false; });
        _updatingArchivePathText     = true;
        const std::wstring text      = path.wstring();
        _archivePathCombo->SetItems({RedSalamander::DxUi::ComboBox::Item{text, text}});
        _archivePathCombo->SetText(text);
        _archivePathUserEdited = markUserEdited;
    }

    void ResetArchivePathForSelectedPacker() noexcept
    {
        const ArchivePackerDefinition& packer = SelectedPacker();
        std::filesystem::path path;
        if (_archivePathUserEdited && _archivePathCombo)
        {
            path = std::filesystem::path(StringUtils::TrimWhitespaceCopy(std::wstring(_archivePathCombo->GetText())));
            if (path.empty())
            {
                path = ResolveUniqueArchivePath(_currentFolder, _suggestedStem, packer.extensionNoDot);
            }
            else
            {
                path = EnsureArchivePathExtension(path, packer.extensionNoDot);
                std::error_code ec;
                if (std::filesystem::exists(path, ec) || ec)
                {
                    const std::filesystem::path parent = path.parent_path().empty() ? _currentFolder : path.parent_path();
                    std::wstring stem                  = SanitizeArchiveSuggestedStem(path.stem().wstring());
                    path                               = ResolveUniqueArchivePath(parent, stem, packer.extensionNoDot);
                }
            }
        }
        else
        {
            path = ResolveUniqueArchivePath(_currentFolder, _suggestedStem, packer.extensionNoDot);
        }

        SetArchivePathText(std::move(path), false);
    }

    void Layout() noexcept
    {
        if (! _root)
        {
            return;
        }

        const auto startedAt     = std::chrono::steady_clock::now();
        const D2D1_RECT_F client = _dxHost.GetClientBoundsDip();
        _root->SetBounds(client);

        constexpr float kMarginDip         = 20.0f;
        constexpr float kGapDip            = 8.0f;
        constexpr float kLabelHeightDip    = 22.0f;
        constexpr float kComboHeightDip    = 32.0f;
        constexpr float kCheckboxHeightDip = 30.0f;
        constexpr float kButtonWidthDip    = 110.0f;
        constexpr float kButtonHeightDip   = 34.0f;

        const float left  = client.left + kMarginDip;
        const float right = std::max(left, client.right - kMarginDip);
        float y           = client.top + kMarginDip;

        if (_promptLabel)
        {
            _promptLabel->SetBounds(D2D1::RectF(left, y, right, y + kLabelHeightDip));
        }
        y += kLabelHeightDip + 2.0f;

        if (_archivePathCombo)
        {
            _archivePathCombo->SetBounds(D2D1::RectF(left, y, right, y + kComboHeightDip));
        }
        y += kComboHeightDip + (kGapDip * 2.0f);

        if (_packerLabel)
        {
            _packerLabel->SetBounds(D2D1::RectF(left, y, right, y + kLabelHeightDip));
        }
        y += kLabelHeightDip + 2.0f;

        if (_packerCombo)
        {
            _packerCombo->SetBounds(D2D1::RectF(left, y, right, y + kComboHeightDip));
        }
        y += kComboHeightDip + kGapDip;

        if (_deleteAfterCheckbox)
        {
            _deleteAfterCheckbox->SetBounds(D2D1::RectF(left, y, right, y + kCheckboxHeightDip));
        }
        y += kCheckboxHeightDip + kGapDip;

        const float totalButtonsWidth = (kButtonWidthDip * 3.0f) + (kGapDip * 2.0f);
        const float buttonsLeft       = std::max(left, left + ((right - left - totalButtonsWidth) * 0.5f));
        const float buttonsTop        = std::max(y, client.bottom - kMarginDip - kButtonHeightDip);
        const float cancelLeft        = buttonsLeft + kButtonWidthDip + kGapDip;
        const float helpLeft          = cancelLeft + kButtonWidthDip + kGapDip;

        if (_okButton)
        {
            _okButton->SetBounds(D2D1::RectF(buttonsLeft, buttonsTop, buttonsLeft + kButtonWidthDip, buttonsTop + kButtonHeightDip));
        }
        if (_cancelButton)
        {
            _cancelButton->SetBounds(D2D1::RectF(cancelLeft, buttonsTop, cancelLeft + kButtonWidthDip, buttonsTop + kButtonHeightDip));
        }
        if (_helpButton)
        {
            _helpButton->SetBounds(D2D1::RectF(helpLeft, buttonsTop, helpLeft + kButtonWidthDip, buttonsTop + kButtonHeightDip));
        }

        _lastButtonsBottomDip = buttonsTop + kButtonHeightDip;
        Debug::Perf::Emit(L"commands.dialog.archivePack_layout_us",
                          L"",
                          Debug::Perf::ElapsedUs(startedAt),
                          static_cast<uint64_t>(_packers.size()),
                          static_cast<uint64_t>(CountVisibleChildWindowsLocal(_hWnd.get())),
                          S_OK);
    }

    [[nodiscard]] ArchivePackPromptResult ReadResultFromUi() const
    {
        ArchivePackPromptResult result{};
        result.packer        = SelectedPacker();
        result.deleteSources = _deleteAfterCheckbox && _deleteAfterCheckbox->IsChecked();
        std::filesystem::path archivePath =
            std::filesystem::path(StringUtils::TrimWhitespaceCopy(_archivePathCombo ? std::wstring(_archivePathCombo->GetText()) : std::wstring{}));
        if (! archivePath.empty())
        {
            archivePath = EnsureArchivePathExtension(archivePath, result.packer.extensionNoDot);
            std::error_code ec;
            if (std::filesystem::exists(archivePath, ec) || ec)
            {
                const std::filesystem::path parent = archivePath.parent_path().empty() ? _currentFolder : archivePath.parent_path();
                archivePath = ResolveUniqueArchivePath(parent, SanitizeArchiveSuggestedStem(archivePath.stem().wstring()), result.packer.extensionNoDot);
            }
        }
        result.archivePath = std::move(archivePath);
        return result;
    }

    void Confirm() noexcept
    {
        _result = ReadResultFromUi();
        _done   = true;
        if (_hWnd && IsWindow(_hWnd.get()) != FALSE)
        {
            _hWnd.reset();
        }
    }

    void Cancel() noexcept
    {
        _result.reset();
        _done = true;
        if (_hWnd && IsWindow(_hWnd.get()) != FALSE)
        {
            _hWnd.reset();
        }
    }

    void SetPackerIndex(size_t index) noexcept
    {
        if (! _packerCombo || index >= _packers.size())
        {
            return;
        }
        _packerCombo->SetSelectedIndex(index);
        ResetArchivePathForSelectedPacker();
    }

    void SetArchivePathForDebug(std::wstring_view path) noexcept
    {
        SetArchivePathText(std::filesystem::path(std::wstring(path)), true);
    }

#ifdef ENABLE_TESTS
    LRESULT OnDebugCommand(ArchivePackPromptDebugCommand command, LPARAM lParam) noexcept
    {
        switch (command)
        {
            case ArchivePackPromptDebugCommand::GetSnapshot:
            {
                auto* snapshot = reinterpret_cast<ArchivePackPromptDebugSnapshot*>(lParam);
                if (! snapshot)
                {
                    return FALSE;
                }

                const ArchivePackerDefinition& selectedPacker = SelectedPacker();
                *snapshot                                     = ArchivePackPromptDebugSnapshot{};
                snapshot->usesDxUiHost                        = _dxHost.GetRoot() != nullptr;
                snapshot->visibleChildWindowCount             = CountVisibleChildWindowsLocal(_hWnd.get());
                snapshot->visibleNativeChildControlCount      = CountVisibleNativeChildControlWindowsLocal(_hWnd.get());
                snapshot->dialogClassName                     = GetWindowClassNameLocal(_hWnd.get());
                snapshot->archivePathText                     = _archivePathCombo ? std::wstring(_archivePathCombo->GetText()) : std::wstring{};
                snapshot->packerDisplayName                   = selectedPacker.displayName;
                snapshot->packerExtension                     = selectedPacker.extensionNoDot;
                snapshot->packerCount                         = _packers.size();
                snapshot->selectedPackerIndex                 = SelectedPackerIndex();
                snapshot->deleteAfterPacking                  = _deleteAfterCheckbox && _deleteAfterCheckbox->IsChecked();
                const D2D1_RECT_F client                      = _dxHost.GetClientBoundsDip();
                snapshot->commandButtonsFitInClient           = _lastButtonsBottomDip <= client.bottom;
                return TRUE;
            }
            case ArchivePackPromptDebugCommand::SetPackerIndex: SetPackerIndex(static_cast<size_t>(lParam)); return TRUE;
            case ArchivePackPromptDebugCommand::SetArchivePath:
            {
                const auto* payload = reinterpret_cast<const ArchivePackPromptArchivePathPayload*>(lParam);
                if (! payload)
                {
                    return FALSE;
                }
                SetArchivePathForDebug(payload->archivePath);
                return TRUE;
            }
            case ArchivePackPromptDebugCommand::SetDeleteAfter:
                if (_deleteAfterCheckbox)
                {
                    _deleteAfterCheckbox->SetChecked(lParam != 0);
                }
                return TRUE;
            case ArchivePackPromptDebugCommand::Confirm: Confirm(); return TRUE;
            case ArchivePackPromptDebugCommand::Cancel: Cancel(); return TRUE;
        }

        return FALSE;
    }
#endif

    [[nodiscard]] static bool ShouldContinueModalLoop(void* context) noexcept
    {
        const auto* self = static_cast<const ArchivePackPromptWindow*>(context);
        return self && ! self->_done;
    }

    static void OnModalLoopQuit(WPARAM, void* context) noexcept
    {
        auto* self = static_cast<ArchivePackPromptWindow*>(context);
        if (self)
        {
            self->_done = true;
        }
    }

private:
    HWND _ownerWindow        = nullptr;
    HWND _restoreFocusWindow = nullptr;
    AppTheme _theme{};
    std::filesystem::path _currentFolder;
    std::vector<std::filesystem::path> _selectedPaths;
    std::vector<ArchivePackerDefinition> _packers;
    std::wstring _suggestedStem;
    wil::unique_hwnd _hWnd;
    RedSalamander::DxUi::WindowHost _dxHost;
    RedSalamander::DxUi::ThemePalette _palette{};
    std::unique_ptr<RedSalamander::DxUi::Panel> _rootStorage;
    RedSalamander::DxUi::Panel* _root                   = nullptr;
    RedSalamander::DxUi::Label* _promptLabel            = nullptr;
    RedSalamander::DxUi::ComboBox* _archivePathCombo    = nullptr;
    RedSalamander::DxUi::Label* _packerLabel            = nullptr;
    RedSalamander::DxUi::ComboBox* _packerCombo         = nullptr;
    RedSalamander::DxUi::Checkbox* _deleteAfterCheckbox = nullptr;
    RedSalamander::DxUi::Button* _okButton              = nullptr;
    RedSalamander::DxUi::Button* _cancelButton          = nullptr;
    RedSalamander::DxUi::Button* _helpButton            = nullptr;
    bool _done                                          = false;
    bool _archivePathUserEdited                         = false;
    bool _updatingArchivePathText                       = false;
    float _lastButtonsBottomDip                         = 0.0f;
    std::optional<ArchivePackPromptResult> _result;
};

[[nodiscard]] std::optional<ArchivePackPromptResult> PromptForArchivePackOptions(HWND ownerWindow,
                                                                                 const AppTheme& theme,
                                                                                 const std::filesystem::path& currentFolder,
                                                                                 const std::vector<std::filesystem::path>& selectedPaths)
{
    ArchivePackPromptWindow prompt(ownerWindow, theme, currentFolder, selectedPaths, EnumerateArchivePackerDefinitions());
    return prompt.ShowModal();
}

class ArchiveUnpackPromptWindow final
{
public:
    ArchiveUnpackPromptWindow(const ArchiveUnpackPromptWindow&)            = delete;
    ArchiveUnpackPromptWindow& operator=(const ArchiveUnpackPromptWindow&) = delete;
    ArchiveUnpackPromptWindow(ArchiveUnpackPromptWindow&&)                 = delete;
    ArchiveUnpackPromptWindow& operator=(ArchiveUnpackPromptWindow&&)      = delete;

    ArchiveUnpackPromptWindow(HWND ownerWindow,
                              const AppTheme& theme,
                              std::filesystem::path currentFolder,
                              std::filesystem::path archivePath,
                              std::vector<ArchiveUnpackerDefinition> unpackers) noexcept
        : _ownerWindow(GetOwnerWindowOrSelf(ownerWindow)),
          _restoreFocusWindow(ownerWindow && IsWindow(ownerWindow) != FALSE ? ownerWindow : nullptr),
          _theme(theme),
          _currentFolder(std::move(currentFolder)),
          _archivePath(std::move(archivePath)),
          _unpackers(std::move(unpackers))
    {
        if (_unpackers.empty())
        {
            _unpackers.push_back(BuildStoredZipUnpackerDefinition());
        }

        _suggestedDestinationPath = BuildArchiveUnpackSuggestedDestinationPath(_archivePath, _currentFolder);
        if (_ownerWindow && IsWindow(_ownerWindow) != FALSE)
        {
            const HWND focused = GetFocus();
            if (focused && IsWindow(focused) != FALSE && (focused == _ownerWindow || IsChild(_ownerWindow, focused) != FALSE))
            {
                _restoreFocusWindow = focused;
            }
            else if (! _restoreFocusWindow || IsWindow(_restoreFocusWindow) == FALSE ||
                     (_restoreFocusWindow != _ownerWindow && IsChild(_ownerWindow, _restoreFocusWindow) == FALSE))
            {
                _restoreFocusWindow = _ownerWindow;
            }
        }
        else if (! _restoreFocusWindow || IsWindow(_restoreFocusWindow) == FALSE)
        {
            _restoreFocusWindow = nullptr;
        }
    }

    [[nodiscard]] std::optional<ArchiveUnpackPromptResult> ShowModal() noexcept
    {
        const HRESULT classHr = EnsureWindowClass();
        if (FAILED(classHr))
        {
            return std::nullopt;
        }

        const DWORD style   = WS_CAPTION | WS_SYSMENU | WS_POPUP | WS_CLIPCHILDREN | WS_CLIPSIBLINGS;
        const DWORD exStyle = WS_EX_DLGMODALFRAME;
        const UINT dpi      = _ownerWindow && IsWindow(_ownerWindow) != FALSE ? GetDpiForWindow(_ownerWindow) : GetDpiForSystem();

        RECT bounds{0, 0, ScalePanePromptForDpi(dpi, 520), ScalePanePromptForDpi(dpi, _maskHelpVisible ? 562 : 448)};
        if (AdjustWindowRectExForDpi(&bounds, style, FALSE, exStyle, dpi) == FALSE)
        {
            return std::nullopt;
        }

        const bool restoreOwnerEnabled = _ownerWindow && IsWindow(_ownerWindow) != FALSE && IsWindowEnabled(_ownerWindow) != FALSE;
        if (restoreOwnerEnabled)
        {
            EnableWindow(_ownerWindow, FALSE);
        }
        const auto restoreOwner = wil::scope_exit([this, restoreOwnerEnabled]() noexcept
        {
            if (restoreOwnerEnabled && _ownerWindow && IsWindow(_ownerWindow) != FALSE)
            {
                EnableWindow(_ownerWindow, TRUE);
                SetActiveWindow(_ownerWindow);
                const HWND restoreFocus = (_restoreFocusWindow && IsWindow(_restoreFocusWindow) != FALSE &&
                                           (_restoreFocusWindow == _ownerWindow || IsChild(_ownerWindow, _restoreFocusWindow) != FALSE))
                                              ? _restoreFocusWindow
                                              : _ownerWindow;
                SetFocus(restoreFocus);
            }
        });

        const std::wstring caption = LoadStringResource(nullptr, IDS_CMD_UNPACK);
        const HWND hwnd            = CreateWindowExW(exStyle,
                                                     kArchiveUnpackPromptClassName,
                                                     caption.c_str(),
                                                     style,
                                                     CW_USEDEFAULT,
                                                     CW_USEDEFAULT,
                                                     bounds.right - bounds.left,
                                                     bounds.bottom - bounds.top,
                                                     _ownerWindow,
                                                     nullptr,
                                                     GetModuleHandleW(nullptr),
                                                     this);
        if (! hwnd)
        {
            return std::nullopt;
        }
        if (! _hWnd)
        {
            _hWnd.reset(hwnd);
        }

        static_cast<void>(Common::WindowSizing::CenterExistingWindowOnOwner(_hWnd.get(), _ownerWindow));
        static_cast<void>(_dxHost.PrimeForShow());
        ShowWindow(_hWnd.get(), SW_SHOWNORMAL);
        UpdateWindow(_hWnd.get());
        SetForegroundWindow(_hWnd.get());

        const RedSalamander::DxUi::DxUiModalLoopResult loopResult = RedSalamander::DxUi::RunDxUiModalLoop(_hWnd.get(),
                                                                                                          RedSalamander::DxUi::DxUiModalLoopOptions{
                                                                                                              .diagnosticName = L"ArchiveUnpackPrompt",
                                                                                                              .shouldContinue = ShouldContinueModalLoop,
                                                                                                              .context        = this,
                                                                                                              .onQuit         = OnModalLoopQuit,
                                                                                                          });
        if (loopResult == RedSalamander::DxUi::DxUiModalLoopResult::GetMessageFailed)
        {
            return std::nullopt;
        }

        return _result;
    }

    static LRESULT CALLBACK WndProc(HWND hwnd, UINT message, WPARAM wParam, LPARAM lParam) noexcept
    {
        if (message == WM_NCCREATE)
        {
            auto* cs   = reinterpret_cast<CREATESTRUCTW*>(lParam);
            auto* self = static_cast<ArchiveUnpackPromptWindow*>(cs ? cs->lpCreateParams : nullptr);
            if (! self)
            {
                return FALSE;
            }

            SetWindowLongPtrW(hwnd, GWLP_USERDATA, reinterpret_cast<LONG_PTR>(self));
            if (! self->_hWnd)
            {
                self->_hWnd.reset(hwnd);
            }
            return TRUE;
        }

        auto* self = reinterpret_cast<ArchiveUnpackPromptWindow*>(GetWindowLongPtrW(hwnd, GWLP_USERDATA));
        if (! self)
        {
            return DefWindowProcW(hwnd, message, wParam, lParam);
        }

        bool handled     = false;
        LRESULT dxResult = 0;
        if (message != WM_CREATE)
        {
            dxResult = self->_dxHost.HandleMessage(hwnd, message, wParam, lParam, handled);
        }
        if (handled)
        {
            if (message == WM_SIZE || message == WM_DPICHANGED)
            {
                self->Layout();
            }
            if (message == WM_NCDESTROY)
            {
#ifdef ENABLE_TESTS
                g_archiveUnpackPromptWindow.store(nullptr);
#endif
                SetWindowLongPtrW(hwnd, GWLP_USERDATA, 0);
                self->_dxHost.Detach();
                if (self->_hWnd.get() == hwnd)
                {
                    static_cast<void>(self->_hWnd.release());
                }
                self->_done = true;
            }
            return dxResult;
        }

#ifdef ENABLE_TESTS
        if (message == GetArchiveUnpackPromptDebugMessage())
        {
            return self->OnDebugCommand(static_cast<ArchiveUnpackPromptDebugCommand>(wParam), lParam);
        }
#endif

        switch (message)
        {
            case WM_CREATE: return self->OnCreate(hwnd) ? 0 : -1;
            case WM_SIZE: self->Layout(); return 0;
            case WM_DPICHANGED:
            {
                const auto* suggested = reinterpret_cast<const RECT*>(lParam);
                if (suggested)
                {
                    SetWindowPos(hwnd,
                                 nullptr,
                                 suggested->left,
                                 suggested->top,
                                 suggested->right - suggested->left,
                                 suggested->bottom - suggested->top,
                                 SWP_NOZORDER | SWP_NOACTIVATE);
                }
                self->Layout();
                return 0;
            }
            case WM_ERASEBKGND: return 1;
            case WM_NCACTIVATE: ApplyTitleBarTheme(hwnd, self->_theme, wParam != FALSE); return DefWindowProcW(hwnd, message, wParam, lParam);
            case WM_CLOSE: self->Cancel(); return 0;
            case WM_NCDESTROY:
#ifdef ENABLE_TESTS
                g_archiveUnpackPromptWindow.store(nullptr);
#endif
                self->_dxHost.Detach();
                if (self->_hWnd.get() == hwnd)
                {
                    self->_hWnd.release();
                }
                SetWindowLongPtrW(hwnd, GWLP_USERDATA, 0);
                self->_done = true;
                break;
        }

        return DefWindowProcW(hwnd, message, wParam, lParam);
    }

private:
    [[nodiscard]] static HRESULT EnsureWindowClass() noexcept
    {
        static ATOM atom = 0;
        if (atom != 0)
        {
            return S_OK;
        }

        WNDCLASSEXW wc{};
        wc.cbSize        = sizeof(wc);
        wc.lpfnWndProc   = ArchiveUnpackPromptWindow::WndProc;
        wc.hInstance     = GetModuleHandleW(nullptr);
        wc.hCursor       = LoadCursorW(nullptr, IDC_ARROW);
        wc.lpszClassName = kArchiveUnpackPromptClassName;
        wc.style         = CS_DBLCLKS;

        atom = RegisterClassExW(&wc);
        return atom != 0 ? S_OK : HRESULT_FROM_WIN32(GetLastError());
    }

    [[nodiscard]] bool OnCreate(HWND hwnd) noexcept
    {
        if (! _dxHost.Attach(hwnd))
        {
            return false;
        }

#ifdef ENABLE_TESTS
        g_archiveUnpackPromptWindow.store(hwnd);
#endif
        BuildUi();
        ApplyTheme();
        SetDestinationPathText(_suggestedDestinationPath);
        Layout();
        if (_destinationCombo)
        {
            _dxHost.SetFocusControl(_destinationCombo);
        }
        _dxHost.SetDefaultButton(_okButton);
        _dxHost.SetCancelButton(_cancelButton);
        return true;
    }

    void BuildUi()
    {
        if (_root)
        {
            return;
        }

        using namespace RedSalamander::DxUi;

        _rootStorage = std::make_unique<Panel>();
        _root        = _rootStorage.get();

        const std::wstring selectedName = ShortDisplayNameForArchiveSelection({_archivePath});
        std::wstring promptText         = FormatStringResource(nullptr, IDS_ARCHIVE_UNPACK_ARCHIVE_FMT, selectedName);
        if (promptText.empty())
        {
            promptText = LoadStringResource(nullptr, IDS_CMD_UNPACK);
        }

        _promptLabel = _root->AddChild<Label>(std::move(promptText));
        _promptLabel->SetAlignment(DWRITE_TEXT_ALIGNMENT_LEADING);

        _destinationCombo = _root->AddChild<ComboBox>();
        _destinationCombo->SetVariant(ComboBoxVariant::Window);
        _destinationCombo->SetEditable(true);
        _destinationCombo->SetOnSubmitted([this] { Confirm(); });

        _unpackerLabel = _root->AddChild<Label>(LoadStringResource(nullptr, IDS_ARCHIVE_UNPACKER_LABEL));
        _unpackerLabel->SetAlignment(DWRITE_TEXT_ALIGNMENT_LEADING);

        _unpackerCombo = _root->AddChild<ComboBox>();
        _unpackerCombo->SetVariant(ComboBoxVariant::Window);
        _unpackerCombo->SetEditable(false);
        _unpackerCombo->SetItems(BuildArchiveUnpackerComboItems(_unpackers));
        _unpackerCombo->SetSelectedIndex(0u);
        _unpackerCombo->SetOnSubmitted([this] { Confirm(); });

        _conflictPolicyLabel = _root->AddChild<Label>(LoadStringResource(nullptr, IDS_ARCHIVE_UNPACK_CONFLICT_LABEL));
        _conflictPolicyLabel->SetAlignment(DWRITE_TEXT_ALIGNMENT_LEADING);

        _conflictPolicyCombo = _root->AddChild<ComboBox>();
        _conflictPolicyCombo->SetVariant(ComboBoxVariant::Window);
        _conflictPolicyCombo->SetEditable(false);
        _conflictPolicyCombo->SetItems(BuildArchiveConflictPolicyComboItems());
        _conflictPolicyCombo->SetSelectedIndex(0u);
        _conflictPolicyCombo->SetOnSubmitted([this] { Confirm(); });

        _deleteAfterCheckbox = _root->AddChild<Checkbox>(LoadStringResource(nullptr, IDS_ARCHIVE_UNPACK_DELETE_AFTER));
        _deleteAfterCheckbox->SetChecked(false);

        _maskLabel = _root->AddChild<Label>(LoadStringResource(nullptr, IDS_ARCHIVE_UNPACK_FILES_LABEL));
        _maskLabel->SetAlignment(DWRITE_TEXT_ALIGNMENT_LEADING);

        _maskField = _root->AddChild<TextField>(L"*.*");
        _maskField->SetMultiline(false);
        _maskField->SetAccessibleName(LoadStringResource(nullptr, IDS_ARCHIVE_UNPACK_FILES_LABEL));
        _maskField->SetOnSubmitted([this] { Confirm(); });

        _maskHintsButton = _root->AddChild<Button>(LoadStringOrFallback(IDS_ARCHIVE_UNPACK_MASK_HINTS, L"mask hints"));
        _maskHintsButton->SetVariant(ButtonVariant::Hyperlink);
        _maskHintsButton->SetOnClick([this]
        {
            _maskHelpVisible = ! _maskHelpVisible;
            UpdateMaskHelpVisibility();
            ResizeForMaskHelpState();
        });

        _maskHelpLabel = _root->AddChild<Label>(LoadStringResource(nullptr, IDS_SELECTION_MASK_HELP_TEXT));
        _maskHelpLabel->SetMultiline(true);
        _maskHelpLabel->SetAlignment(DWRITE_TEXT_ALIGNMENT_LEADING);
        _maskHelpLabel->SetVisible(false);

        _okButton = _root->AddChild<Button>(LoadStringResource(nullptr, IDS_BTN_OK));
        _okButton->SetPrimary(true);
        _okButton->SetOnClick([this] { Confirm(); });

        _cancelButton = _root->AddChild<Button>(LoadStringResource(nullptr, IDS_BTN_CANCEL));
        _cancelButton->SetOnClick([this] { Cancel(); });

        _helpButton = _root->AddChild<Button>(LoadStringResource(nullptr, IDS_BTN_HELP));
        _helpButton->SetOnClick([this]
        {
            _maskHelpVisible = true;
            UpdateMaskHelpVisibility();
            ResizeForMaskHelpState();
        });

        _dxHost.SetRoot(std::move(_rootStorage));
    }

    void ApplyTheme() noexcept
    {
        _palette = MakeAppThemeDxPalette(_theme, _theme.windowBackground);
        _dxHost.SetTheme(_palette);
        if (_hWnd)
        {
            ApplyWindowChromeTheme(_hWnd.get(), _theme, WindowBackdropTarget::Tool, GetActiveWindow() == _hWnd.get());
            static_cast<void>(RedrawWindow(_hWnd.get(), nullptr, nullptr, RDW_FRAME | RDW_INVALIDATE | RDW_UPDATENOW));
        }
    }

    [[nodiscard]] size_t SelectedUnpackerIndex() const noexcept
    {
        const std::optional<size_t> selectedIndex = _unpackerCombo ? _unpackerCombo->GetSelectedIndex() : std::optional<size_t>{0u};
        const size_t index                        = selectedIndex.value_or(0u);
        return index < _unpackers.size() ? index : 0u;
    }

    [[nodiscard]] const ArchiveUnpackerDefinition& SelectedUnpacker() const noexcept
    {
        return _unpackers[SelectedUnpackerIndex()];
    }

    void SetDestinationPathText(const std::filesystem::path& path) noexcept
    {
        if (! _destinationCombo)
        {
            return;
        }

        const std::wstring text = path.wstring();
        _destinationCombo->SetItems({RedSalamander::DxUi::ComboBox::Item{text, text}});
        _destinationCombo->SetText(text);
    }

    void SetMaskText(std::wstring_view mask) noexcept
    {
        if (_maskField)
        {
            _maskField->SetText(std::wstring(mask));
        }
    }

    void UpdateMaskHelpVisibility() noexcept
    {
        if (_maskHelpLabel)
        {
            _maskHelpLabel->SetVisible(_maskHelpVisible);
        }
    }

    void ResizeForMaskHelpState() noexcept
    {
        if (! _hWnd || IsWindow(_hWnd.get()) == FALSE)
        {
            return;
        }

        RECT windowRect{};
        RECT clientRect{};
        if (GetWindowRect(_hWnd.get(), &windowRect) == FALSE || GetClientRect(_hWnd.get(), &clientRect) == FALSE)
        {
            return;
        }

        const UINT dpi                = GetDpiForWindow(_hWnd.get());
        const int currentClientHeight = std::max(0l, clientRect.bottom - clientRect.top);
        const int nonClientHeight     = std::max<int>(0, static_cast<int>((windowRect.bottom - windowRect.top) - currentClientHeight));
        const int targetClientHeight  = ScalePanePromptForDpi(dpi, _maskHelpVisible ? 562 : 448);
        const int targetWindowHeight  = targetClientHeight + nonClientHeight;

        SetWindowPos(_hWnd.get(), nullptr, 0, 0, windowRect.right - windowRect.left, targetWindowHeight, SWP_NOMOVE | SWP_NOZORDER | SWP_NOACTIVATE);
        Layout();
    }

    void Layout() noexcept
    {
        if (! _root)
        {
            return;
        }

        const auto startedAt     = std::chrono::steady_clock::now();
        const D2D1_RECT_F client = _dxHost.GetClientBoundsDip();
        _root->SetBounds(client);

        constexpr float kMarginDip         = 20.0f;
        constexpr float kGapDip            = 8.0f;
        constexpr float kLabelHeightDip    = 22.0f;
        constexpr float kComboHeightDip    = 32.0f;
        constexpr float kCheckboxHeightDip = 30.0f;
        constexpr float kMaskLinkWidthDip  = 110.0f;
        constexpr float kHelpHeightDip     = 120.0f;
        constexpr float kButtonWidthDip    = 110.0f;
        constexpr float kButtonHeightDip   = 34.0f;

        const float left  = client.left + kMarginDip;
        const float right = std::max(left, client.right - kMarginDip);
        float y           = client.top + kMarginDip;

        if (_promptLabel)
        {
            _promptLabel->SetBounds(D2D1::RectF(left, y, right, y + kLabelHeightDip));
        }
        y += kLabelHeightDip + 2.0f;

        if (_destinationCombo)
        {
            _destinationCombo->SetBounds(D2D1::RectF(left, y, right, y + kComboHeightDip));
        }
        y += kComboHeightDip + (kGapDip * 2.0f);

        if (_unpackerLabel)
        {
            _unpackerLabel->SetBounds(D2D1::RectF(left, y, right, y + kLabelHeightDip));
        }
        y += kLabelHeightDip + 2.0f;

        if (_unpackerCombo)
        {
            _unpackerCombo->SetBounds(D2D1::RectF(left, y, right, y + kComboHeightDip));
        }
        y += kComboHeightDip + (kGapDip * 2.0f);

        if (_conflictPolicyLabel)
        {
            _conflictPolicyLabel->SetBounds(D2D1::RectF(left, y, right, y + kLabelHeightDip));
        }
        y += kLabelHeightDip + 2.0f;

        if (_conflictPolicyCombo)
        {
            _conflictPolicyCombo->SetBounds(D2D1::RectF(left, y, right, y + kComboHeightDip));
        }
        y += kComboHeightDip + kGapDip;

        if (_deleteAfterCheckbox)
        {
            _deleteAfterCheckbox->SetBounds(D2D1::RectF(left, y, right, y + kCheckboxHeightDip));
        }
        y += kCheckboxHeightDip + (kGapDip * 2.0f);

        if (_maskLabel)
        {
            _maskLabel->SetBounds(D2D1::RectF(left, y, right, y + kLabelHeightDip));
        }
        y += kLabelHeightDip + 2.0f;

        if (_maskField)
        {
            _maskField->SetBounds(D2D1::RectF(left, y, right, y + kComboHeightDip));
        }
        y += kComboHeightDip + 2.0f;

        if (_maskHintsButton)
        {
            _maskHintsButton->SetBounds(D2D1::RectF(std::max(left, right - kMaskLinkWidthDip), y, right, y + kLabelHeightDip));
        }
        y += kLabelHeightDip + kGapDip;

        if (_maskHelpLabel)
        {
            _maskHelpLabel->SetBounds(D2D1::RectF(left, y, right, y + (_maskHelpVisible ? kHelpHeightDip : 0.0f)));
        }
        if (_maskHelpVisible)
        {
            y += kHelpHeightDip + kGapDip;
        }

        const float totalButtonsWidth = (kButtonWidthDip * 3.0f) + (kGapDip * 2.0f);
        const float buttonsLeft       = std::max(left, left + ((right - left - totalButtonsWidth) * 0.5f));
        const float buttonsTop        = std::max(y, client.bottom - kMarginDip - kButtonHeightDip);
        const float cancelLeft        = buttonsLeft + kButtonWidthDip + kGapDip;
        const float helpLeft          = cancelLeft + kButtonWidthDip + kGapDip;

        if (_okButton)
        {
            _okButton->SetBounds(D2D1::RectF(buttonsLeft, buttonsTop, buttonsLeft + kButtonWidthDip, buttonsTop + kButtonHeightDip));
        }
        if (_cancelButton)
        {
            _cancelButton->SetBounds(D2D1::RectF(cancelLeft, buttonsTop, cancelLeft + kButtonWidthDip, buttonsTop + kButtonHeightDip));
        }
        if (_helpButton)
        {
            _helpButton->SetBounds(D2D1::RectF(helpLeft, buttonsTop, helpLeft + kButtonWidthDip, buttonsTop + kButtonHeightDip));
        }

        _lastButtonsBottomDip = buttonsTop + kButtonHeightDip;
        Debug::Perf::Emit(L"commands.dialog.archiveUnpack_layout_us",
                          L"",
                          Debug::Perf::ElapsedUs(startedAt),
                          static_cast<uint64_t>(_unpackers.size()),
                          static_cast<uint64_t>(CountVisibleChildWindowsLocal(_hWnd.get())),
                          S_OK);
    }

    [[nodiscard]] ArchiveUnpackPromptResult ReadResultFromUi() const
    {
        ArchiveUnpackPromptResult result{};
        result.unpacker       = SelectedUnpacker();
        result.conflictPolicy = _conflictPolicyCombo && _conflictPolicyCombo->GetSelectedIndex().value_or(0u) == 1u ? ArchiveExistingTargetPolicy::Replace
                                                                                                                    : ArchiveExistingTargetPolicy::Skip;
        result.deleteArchive  = _deleteAfterCheckbox && _deleteAfterCheckbox->IsChecked();
        result.destinationPath =
            std::filesystem::path(StringUtils::TrimWhitespaceCopy(_destinationCombo ? std::wstring(_destinationCombo->GetText()) : std::wstring{}));
        result.maskText = StringUtils::TrimWhitespaceCopy(_maskField ? std::wstring(_maskField->GetText()) : std::wstring{});
        if (result.maskText.empty())
        {
            result.maskText = L"*.*";
        }
        return result;
    }

    void Confirm() noexcept
    {
        _result = ReadResultFromUi();
        _done   = true;
        if (_hWnd && IsWindow(_hWnd.get()) != FALSE)
        {
            _hWnd.reset();
        }
    }

    void Cancel() noexcept
    {
        _result.reset();
        _done = true;
        if (_hWnd && IsWindow(_hWnd.get()) != FALSE)
        {
            _hWnd.reset();
        }
    }

#ifdef ENABLE_TESTS
    LRESULT OnDebugCommand(ArchiveUnpackPromptDebugCommand command, LPARAM lParam) noexcept
    {
        switch (command)
        {
            case ArchiveUnpackPromptDebugCommand::GetSnapshot:
            {
                auto* snapshot = reinterpret_cast<ArchiveUnpackPromptDebugSnapshot*>(lParam);
                if (! snapshot)
                {
                    return FALSE;
                }

                const ArchiveUnpackerDefinition& selectedUnpacker = SelectedUnpacker();
                *snapshot                                         = ArchiveUnpackPromptDebugSnapshot{};
                snapshot->usesDxUiHost                            = _dxHost.GetRoot() != nullptr;
                snapshot->visibleChildWindowCount                 = CountVisibleChildWindowsLocal(_hWnd.get());
                snapshot->visibleNativeChildControlCount          = CountVisibleNativeChildControlWindowsLocal(_hWnd.get());
                snapshot->dialogClassName                         = GetWindowClassNameLocal(_hWnd.get());
                snapshot->destinationPathText                     = _destinationCombo ? std::wstring(_destinationCombo->GetText()) : std::wstring{};
                snapshot->unpackerDisplayName                     = selectedUnpacker.displayName;
                snapshot->unpackerExtension                       = selectedUnpacker.extensionNoDot;
                snapshot->unpackerCount                           = _unpackers.size();
                snapshot->selectedUnpackerIndex                   = SelectedUnpackerIndex();
                snapshot->conflictPolicyIndex                     = _conflictPolicyCombo ? _conflictPolicyCombo->GetSelectedIndex().value_or(0u) : 0u;
                snapshot->replaceExistingFiles                    = snapshot->conflictPolicyIndex == 1u;
                snapshot->deleteAfterUnpacking                    = _deleteAfterCheckbox && _deleteAfterCheckbox->IsChecked();
                snapshot->maskText                                = _maskField ? std::wstring(_maskField->GetText()) : std::wstring{};
                snapshot->maskHelpVisible                         = _maskHelpVisible;
                const D2D1_RECT_F client                          = _dxHost.GetClientBoundsDip();
                snapshot->commandButtonsFitInClient               = _lastButtonsBottomDip <= client.bottom;
                return TRUE;
            }
            case ArchiveUnpackPromptDebugCommand::SetDestinationPath:
            {
                const auto* payload = reinterpret_cast<const ArchiveUnpackPromptTextPayload*>(lParam);
                if (! payload)
                {
                    return FALSE;
                }
                SetDestinationPathText(std::filesystem::path(payload->text));
                return TRUE;
            }
            case ArchiveUnpackPromptDebugCommand::SetMask:
            {
                const auto* payload = reinterpret_cast<const ArchiveUnpackPromptTextPayload*>(lParam);
                if (! payload)
                {
                    return FALSE;
                }
                SetMaskText(payload->text);
                return TRUE;
            }
            case ArchiveUnpackPromptDebugCommand::SetConflictPolicy:
                if (_conflictPolicyCombo)
                {
                    _conflictPolicyCombo->SetSelectedIndex(lParam != 0 ? 1u : 0u);
                }
                return TRUE;
            case ArchiveUnpackPromptDebugCommand::SetDeleteAfter:
                if (_deleteAfterCheckbox)
                {
                    _deleteAfterCheckbox->SetChecked(lParam != 0);
                }
                return TRUE;
            case ArchiveUnpackPromptDebugCommand::Confirm: Confirm(); return TRUE;
            case ArchiveUnpackPromptDebugCommand::Cancel: Cancel(); return TRUE;
        }

        return FALSE;
    }
#endif

    [[nodiscard]] static bool ShouldContinueModalLoop(void* context) noexcept
    {
        const auto* self = static_cast<const ArchiveUnpackPromptWindow*>(context);
        return self && ! self->_done;
    }

    static void OnModalLoopQuit(WPARAM, void* context) noexcept
    {
        auto* self = static_cast<ArchiveUnpackPromptWindow*>(context);
        if (self)
        {
            self->_done = true;
        }
    }

private:
    HWND _ownerWindow        = nullptr;
    HWND _restoreFocusWindow = nullptr;
    AppTheme _theme{};
    std::filesystem::path _currentFolder;
    std::filesystem::path _archivePath;
    std::filesystem::path _suggestedDestinationPath;
    std::vector<ArchiveUnpackerDefinition> _unpackers;
    wil::unique_hwnd _hWnd;
    RedSalamander::DxUi::WindowHost _dxHost;
    RedSalamander::DxUi::ThemePalette _palette{};
    std::unique_ptr<RedSalamander::DxUi::Panel> _rootStorage;
    RedSalamander::DxUi::Panel* _root                   = nullptr;
    RedSalamander::DxUi::Label* _promptLabel            = nullptr;
    RedSalamander::DxUi::ComboBox* _destinationCombo    = nullptr;
    RedSalamander::DxUi::Label* _unpackerLabel          = nullptr;
    RedSalamander::DxUi::ComboBox* _unpackerCombo       = nullptr;
    RedSalamander::DxUi::Label* _conflictPolicyLabel    = nullptr;
    RedSalamander::DxUi::ComboBox* _conflictPolicyCombo = nullptr;
    RedSalamander::DxUi::Checkbox* _deleteAfterCheckbox = nullptr;
    RedSalamander::DxUi::Label* _maskLabel              = nullptr;
    RedSalamander::DxUi::TextField* _maskField          = nullptr;
    RedSalamander::DxUi::Button* _maskHintsButton       = nullptr;
    RedSalamander::DxUi::Label* _maskHelpLabel          = nullptr;
    RedSalamander::DxUi::Button* _okButton              = nullptr;
    RedSalamander::DxUi::Button* _cancelButton          = nullptr;
    RedSalamander::DxUi::Button* _helpButton            = nullptr;
    bool _done                                          = false;
    bool _maskHelpVisible                               = false;
    float _lastButtonsBottomDip                         = 0.0f;
    std::optional<ArchiveUnpackPromptResult> _result;
};

[[nodiscard]] std::optional<ArchiveUnpackPromptResult> PromptForArchiveUnpackOptions(HWND ownerWindow,
                                                                                     const AppTheme& theme,
                                                                                     const std::filesystem::path& currentFolder,
                                                                                     const std::filesystem::path& archivePath)
{
    ArchiveUnpackPromptWindow prompt(ownerWindow, theme, currentFolder, archivePath, {BuildStoredZipUnpackerDefinition()});
    return prompt.ShowModal();
}

constexpr uint32_t kZipLocalFileHeaderSignature   = 0x04034B50u;
constexpr uint32_t kZipCentralDirectorySignature  = 0x02014B50u;
constexpr uint32_t kZipEndOfCentralDirSignature   = 0x06054B50u;
constexpr uint16_t kZipVersionStored              = 20u;
constexpr uint16_t kZipGeneralPurposeUtf8         = 0x0800u;
constexpr uint16_t kZipCompressionStored          = 0u;
constexpr uint16_t kZipDosTimeMidnight            = 0u;
constexpr uint16_t kZipDosDate1980Jan1            = 33u;
constexpr uint16_t kZipDirectoryExternalAttribute = 0x0010u;
constexpr size_t kZipEndOfCentralDirMinSize       = 22u;
constexpr size_t kZipMaxCommentSize               = 0xFFFFu;
constexpr size_t kZipMaxTailSearchSize            = kZipEndOfCentralDirMinSize + kZipMaxCommentSize;
constexpr uint64_t kMaxArchiveExtractEntryBytes   = 4ull * 1024ull * 1024ull * 1024ull;
constexpr uint64_t kMaxArchiveExtractTotalBytes   = 8ull * 1024ull * 1024ull * 1024ull;

[[nodiscard]] HRESULT AccumulateArchiveExtractSize(uint64_t& totalBytes, uint64_t entryBytes) noexcept
{
    if (entryBytes > kMaxArchiveExtractEntryBytes || totalBytes > kMaxArchiveExtractTotalBytes || entryBytes > kMaxArchiveExtractTotalBytes - totalBytes)
    {
        return HRESULT_FROM_WIN32(ERROR_FILE_TOO_LARGE);
    }

    totalBytes += entryBytes;
    return S_OK;
}

struct ArchiveEntrySource final
{
    std::filesystem::path sourcePath;
    std::wstring entryName;
    bool directory    = false;
    uint64_t fileSize = 0u;
};

struct ZipCentralEntry final
{
    std::wstring entryName;
    std::string entryNameUtf8;
    uint32_t crc32             = 0u;
    uint32_t compressedSize    = 0u;
    uint32_t uncompressedSize  = 0u;
    uint32_t localHeaderOffset = 0u;
    bool directory             = false;
};

struct ZipParsedEntry final
{
    std::wstring entryName;
    std::filesystem::path relativePath;
    uint32_t crc32             = 0u;
    uint32_t compressedSize    = 0u;
    uint32_t uncompressedSize  = 0u;
    uint32_t localHeaderOffset = 0u;
    bool directory             = false;
};

struct ZipEndOfCentralDirectory final
{
    uint16_t entryCount             = 0u;
    uint32_t centralDirectorySize   = 0u;
    uint32_t centralDirectoryOffset = 0u;
};

struct ArchiveOperationResult final
{
    HRESULT hr = S_OK;
    std::filesystem::path archivePath;
    std::filesystem::path destinationPath;
    uint64_t entryCount           = 0u;
    uint64_t bytesProcessed       = 0u;
    uint64_t skippedConflictCount = 0u;
    std::vector<std::wstring> entries;
};

void AppendLe16(std::vector<std::byte>& bytes, uint16_t value)
{
    bytes.push_back(static_cast<std::byte>(value & 0xFFu));
    bytes.push_back(static_cast<std::byte>((value >> 8u) & 0xFFu));
}

void AppendLe32(std::vector<std::byte>& bytes, uint32_t value)
{
    AppendLe16(bytes, static_cast<uint16_t>(value & 0xFFFFu));
    AppendLe16(bytes, static_cast<uint16_t>((value >> 16u) & 0xFFFFu));
}

[[nodiscard]] bool TryReadLe16(std::span<const std::byte> bytes, size_t offset, uint16_t& out) noexcept
{
    if (offset > bytes.size() || bytes.size() - offset < sizeof(uint16_t))
    {
        return false;
    }

    out = static_cast<uint16_t>(std::to_integer<uint16_t>(bytes[offset]) | static_cast<uint16_t>(std::to_integer<uint16_t>(bytes[offset + 1u]) << 8u));
    return true;
}

[[nodiscard]] bool TryReadLe32(std::span<const std::byte> bytes, size_t offset, uint32_t& out) noexcept
{
    uint16_t lo = 0u;
    uint16_t hi = 0u;
    if (! TryReadLe16(bytes, offset, lo) || ! TryReadLe16(bytes, offset + sizeof(uint16_t), hi))
    {
        return false;
    }

    out = static_cast<uint32_t>(lo) | (static_cast<uint32_t>(hi) << 16u);
    return true;
}

[[nodiscard]] HRESULT HResultFromErrorCode(const std::error_code& ec) noexcept
{
    if (! ec)
    {
        return S_OK;
    }
    return HRESULT_FROM_WIN32(static_cast<DWORD>(ec.value()));
}

[[nodiscard]] HRESULT HResultFromLastError() noexcept
{
    const DWORD error = GetLastError();
    return HRESULT_FROM_WIN32(error == ERROR_SUCCESS ? ERROR_GEN_FAILURE : error);
}

enum class ArchiveTargetDecision : uint8_t
{
    Extract,
    Skip,
};

[[nodiscard]] Common::Files::ExistingTargetPolicy LocalFilePolicyForArchive(ArchiveExistingTargetPolicy policy) noexcept
{
    return policy == ArchiveExistingTargetPolicy::Replace ? Common::Files::ExistingTargetPolicy::Replace : Common::Files::ExistingTargetPolicy::FailIfExists;
}

[[nodiscard]] bool IsArchiveTargetExistsFailure(HRESULT hr) noexcept
{
    return hr == HRESULT_FROM_WIN32(ERROR_FILE_EXISTS) || hr == HRESULT_FROM_WIN32(ERROR_ALREADY_EXISTS);
}

[[nodiscard]] HRESULT ClassifyArchiveTarget(const std::filesystem::path& targetPath,
                                            bool directory,
                                            ArchiveExistingTargetPolicy policy,
                                            ArchiveTargetDecision& outDecision) noexcept
{
    outDecision            = ArchiveTargetDecision::Extract;
    const DWORD attributes = GetFileAttributesW(targetPath.c_str());
    if (attributes == INVALID_FILE_ATTRIBUTES)
    {
        const DWORD error = GetLastError();
        if (error == ERROR_FILE_NOT_FOUND || error == ERROR_PATH_NOT_FOUND)
        {
            return S_OK;
        }
        return HRESULT_FROM_WIN32(error == ERROR_SUCCESS ? ERROR_GEN_FAILURE : error);
    }

    if (policy == ArchiveExistingTargetPolicy::Skip)
    {
        outDecision = ArchiveTargetDecision::Skip;
        return S_OK;
    }

    const bool existingIsDirectory = (attributes & FILE_ATTRIBUTE_DIRECTORY) != 0u;
    return existingIsDirectory == directory ? S_OK : HRESULT_FROM_WIN32(ERROR_ALREADY_EXISTS);
}

[[nodiscard]] const std::array<uint32_t, 256>& ArchiveCrc32Table() noexcept
{
    static const std::array<uint32_t, 256> table = []
    {
        std::array<uint32_t, 256> values{};
        for (uint32_t i = 0u; i < static_cast<uint32_t>(values.size()); ++i)
        {
            uint32_t crc = i;
            for (uint32_t bit = 0u; bit < 8u; ++bit)
            {
                crc = (crc & 1u) != 0u ? (0xEDB88320u ^ (crc >> 1u)) : (crc >> 1u);
            }
            values[static_cast<size_t>(i)] = crc;
        }
        return values;
    }();
    return table;
}

[[nodiscard]] uint32_t ArchiveCrc32Update(uint32_t crc, const std::byte* data, size_t size) noexcept
{
    const auto& table = ArchiveCrc32Table();
    uint32_t current  = crc;
    for (size_t index = 0u; index < size; ++index)
    {
        const uint32_t byte = std::to_integer<uint32_t>(data[index]);
        current             = table[static_cast<size_t>((current ^ byte) & 0xFFu)] ^ (current >> 8u);
    }
    return current;
}

[[nodiscard]] HRESULT GetArchiveFilePosition(HANDLE file, uint64_t& outPosition) noexcept
{
    outPosition = 0u;
    LARGE_INTEGER zero{};
    LARGE_INTEGER current{};
    if (SetFilePointerEx(file, zero, &current, FILE_CURRENT) == FALSE)
    {
        return HResultFromLastError();
    }
    if (current.QuadPart < 0)
    {
        return HRESULT_FROM_WIN32(ERROR_NEGATIVE_SEEK);
    }

    outPosition = static_cast<uint64_t>(current.QuadPart);
    return S_OK;
}

[[nodiscard]] HRESULT SetArchiveFilePosition(HANDLE file, uint64_t position) noexcept
{
    if (position > static_cast<uint64_t>((std::numeric_limits<LONGLONG>::max)()))
    {
        return HRESULT_FROM_WIN32(ERROR_FILE_TOO_LARGE);
    }

    LARGE_INTEGER target{};
    target.QuadPart = static_cast<LONGLONG>(position);
    return SetFilePointerEx(file, target, nullptr, FILE_BEGIN) == FALSE ? HResultFromLastError() : S_OK;
}

[[nodiscard]] HRESULT WriteArchiveBytes(HANDLE file, const void* data, size_t size) noexcept
{
    const auto* cursor = static_cast<const std::byte*>(data);
    size_t remaining   = size;
    while (remaining > 0u)
    {
        const DWORD chunk = static_cast<DWORD>((std::min)(remaining, static_cast<size_t>((std::numeric_limits<DWORD>::max)())));
        DWORD written     = 0u;
        if (WriteFile(file, cursor, chunk, &written, nullptr) == FALSE)
        {
            return HResultFromLastError();
        }
        if (written == 0u || written > chunk)
        {
            return HRESULT_FROM_WIN32(ERROR_WRITE_FAULT);
        }
        cursor += written;
        remaining -= written;
    }
    return S_OK;
}

[[nodiscard]] HRESULT WriteArchiveBytes(HANDLE file, const std::vector<std::byte>& bytes) noexcept
{
    return bytes.empty() ? S_OK : WriteArchiveBytes(file, bytes.data(), bytes.size());
}

[[nodiscard]] HRESULT ReadArchiveBytesAt(HANDLE file, uint64_t offset, std::span<std::byte> bytes) noexcept
{
    if (const HRESULT hr = SetArchiveFilePosition(file, offset); FAILED(hr))
    {
        return hr;
    }

    size_t remaining = bytes.size();
    size_t written   = 0u;
    while (remaining > 0u)
    {
        const DWORD chunk = static_cast<DWORD>((std::min)(remaining, static_cast<size_t>((std::numeric_limits<DWORD>::max)())));
        DWORD bytesRead   = 0u;
        if (ReadFile(file, bytes.data() + written, chunk, &bytesRead, nullptr) == FALSE)
        {
            return HResultFromLastError();
        }
        if (bytesRead == 0u || bytesRead > chunk)
        {
            return HRESULT_FROM_WIN32(ERROR_HANDLE_EOF);
        }
        written += bytesRead;
        remaining -= bytesRead;
    }
    return S_OK;
}

[[nodiscard]] HRESULT ReadArchiveBytesAt(HANDLE file, uint64_t offset, size_t size, std::vector<std::byte>& outBytes)
{
    outBytes.assign(size, std::byte{});
    return outBytes.empty() ? S_OK : ReadArchiveBytesAt(file, offset, std::span<std::byte>(outBytes.data(), outBytes.size()));
}

[[nodiscard]] std::wstring ArchiveEntryNameFromRelativePath(const std::filesystem::path& relativePath, bool directory)
{
    std::wstring name;
    for (const std::filesystem::path& componentPath : relativePath)
    {
        const std::wstring component = componentPath.native();
        if (component.empty() || component == L"." || component == L"..")
        {
            return {};
        }
        if (! name.empty())
        {
            name.push_back(L'/');
        }
        name.append(component);
    }

    if (name.empty())
    {
        return {};
    }
    if (directory && name.back() != L'/')
    {
        name.push_back(L'/');
    }
    return name;
}

[[nodiscard]] bool IsReservedDosDeviceNameComponent(std::wstring_view component) noexcept
{
    size_t end = component.size();
    while (end > 0u && (component[end - 1u] == L'.' || component[end - 1u] == L' '))
    {
        --end;
    }
    if (end == 0u)
    {
        return false;
    }

    std::wstring_view stem = component.substr(0u, end);
    if (const size_t dot = stem.find(L'.'); dot != std::wstring_view::npos)
    {
        stem = stem.substr(0u, dot);
    }
    if (stem.empty())
    {
        return false;
    }

    if (DirectoryNamesMatch(stem, L"CON", true) || DirectoryNamesMatch(stem, L"PRN", true) || DirectoryNamesMatch(stem, L"AUX", true) ||
        DirectoryNamesMatch(stem, L"NUL", true) || DirectoryNamesMatch(stem, L"CONIN$", true) || DirectoryNamesMatch(stem, L"CONOUT$", true))
    {
        return true;
    }

    if (stem.size() == 4u)
    {
        const wchar_t suffix = stem[3];
        if (suffix >= L'1' && suffix <= L'9')
        {
            return DirectoryNamesMatch(stem.substr(0u, 3u), L"COM", true) || DirectoryNamesMatch(stem.substr(0u, 3u), L"LPT", true);
        }
    }

    return false;
}

[[nodiscard]] bool IsArchiveTargetPathInsideDestination(const std::filesystem::path& destinationPath, const std::filesystem::path& targetPath) noexcept
{
    std::error_code ec;
    const std::filesystem::path destination = std::filesystem::absolute(destinationPath, ec).lexically_normal();
    if (ec || destination.empty())
    {
        return false;
    }

    ec.clear();
    const std::filesystem::path target = std::filesystem::absolute(targetPath, ec).lexically_normal();
    if (ec || target.empty())
    {
        return false;
    }

    const std::filesystem::path relative = target.lexically_relative(destination);
    if (relative.empty())
    {
        return false;
    }

    for (const std::filesystem::path& component : relative)
    {
        if (component == L"..")
        {
            return false;
        }
    }
    return true;
}

[[nodiscard]] bool IsArchiveEntryNameSafe(std::wstring_view entryName, bool& outDirectory, std::filesystem::path& outRelativePath) noexcept
{
    outDirectory = false;
    outRelativePath.clear();
    if (entryName.empty() || entryName.front() == L'/' || entryName.front() == L'\\')
    {
        return false;
    }

    outDirectory = entryName.back() == L'/';
    std::wstring component;
    bool hadComponent = false;
    for (const wchar_t ch : entryName)
    {
        if (ch == L'\\' || ch == L':' || ch == L'\0')
        {
            return false;
        }

        if (ch == L'/')
        {
            if (component.empty() || component == L"." || component == L".." || IsReservedDosDeviceNameComponent(component))
            {
                return false;
            }
            outRelativePath /= component;
            component.clear();
            hadComponent = true;
            continue;
        }

        component.push_back(ch);
    }

    if (! component.empty())
    {
        if (component == L"." || component == L".." || IsReservedDosDeviceNameComponent(component))
        {
            return false;
        }
        outRelativePath /= component;
        hadComponent = true;
    }
    else if (! outDirectory)
    {
        return false;
    }

    return hadComponent && ! outRelativePath.empty();
}

[[nodiscard]] HRESULT AddArchiveSourceEntry(const std::filesystem::path& path,
                                            const std::filesystem::path& relativePath,
                                            bool directory,
                                            std::vector<ArchiveEntrySource>& entries)
{
    ArchiveEntrySource entry{};
    entry.sourcePath = path;
    entry.entryName  = ArchiveEntryNameFromRelativePath(relativePath, directory);
    entry.directory  = directory;
    if (entry.entryName.empty())
    {
        return HRESULT_FROM_WIN32(ERROR_INVALID_NAME);
    }

    if (! directory)
    {
        std::error_code ec;
        entry.fileSize = std::filesystem::file_size(path, ec);
        if (ec)
        {
            return HResultFromErrorCode(ec);
        }
        if (entry.fileSize > static_cast<uint64_t>((std::numeric_limits<uint32_t>::max)()))
        {
            return HRESULT_FROM_WIN32(ERROR_FILE_TOO_LARGE);
        }
    }

    entries.push_back(std::move(entry));
    return S_OK;
}

[[nodiscard]] HRESULT CollectArchiveDirectoryEntries(const std::filesystem::path& root,
                                                     const std::filesystem::path& relativeRoot,
                                                     std::vector<ArchiveEntrySource>& entries)
{
    if (const HRESULT hr = AddArchiveSourceEntry(root, relativeRoot, true, entries); FAILED(hr))
    {
        return hr;
    }

    std::error_code ec;
    std::filesystem::recursive_directory_iterator it(root, std::filesystem::directory_options::skip_permission_denied, ec);
    const std::filesystem::recursive_directory_iterator end;
    if (ec)
    {
        return HResultFromErrorCode(ec);
    }

    const std::filesystem::path parent = root.parent_path();
    for (; it != end;)
    {
        const std::filesystem::path itemPath      = it->path();
        const std::filesystem::file_status status = it->symlink_status(ec);
        if (ec)
        {
            return HResultFromErrorCode(ec);
        }

        const bool isDirectory   = std::filesystem::is_directory(status);
        const bool isRegularFile = std::filesystem::is_regular_file(status);
        if (isDirectory || isRegularFile)
        {
            const std::filesystem::path relativePath = itemPath.lexically_relative(parent);
            if (const HRESULT hr = AddArchiveSourceEntry(itemPath, relativePath, isDirectory, entries); FAILED(hr))
            {
                return hr;
            }
        }

        it.increment(ec);
        if (ec)
        {
            return HResultFromErrorCode(ec);
        }
    }

    return S_OK;
}

[[nodiscard]] HRESULT CollectArchiveSources(const std::vector<std::filesystem::path>& selectedPaths, std::vector<ArchiveEntrySource>& outEntries)
{
    outEntries.clear();
    for (const std::filesystem::path& path : selectedPaths)
    {
        std::error_code ec;
        const std::filesystem::file_status status = std::filesystem::symlink_status(path, ec);
        if (ec)
        {
            return HResultFromErrorCode(ec);
        }
        if (! std::filesystem::exists(status))
        {
            return HRESULT_FROM_WIN32(ERROR_PATH_NOT_FOUND);
        }

        const std::filesystem::path relativeRoot = path.filename();
        if (relativeRoot.empty())
        {
            return HRESULT_FROM_WIN32(ERROR_INVALID_NAME);
        }

        if (std::filesystem::is_directory(status))
        {
            if (const HRESULT hr = CollectArchiveDirectoryEntries(path, relativeRoot, outEntries); FAILED(hr))
            {
                return hr;
            }
        }
        else if (std::filesystem::is_regular_file(status))
        {
            if (const HRESULT hr = AddArchiveSourceEntry(path, relativeRoot, false, outEntries); FAILED(hr))
            {
                return hr;
            }
        }
    }

    std::sort(outEntries.begin(), outEntries.end(), [](const ArchiveEntrySource& left, const ArchiveEntrySource& right) noexcept {
        return left.entryName < right.entryName;
    });
    outEntries.erase(std::unique(outEntries.begin(),
                                 outEntries.end(),
                                 [](const ArchiveEntrySource& left, const ArchiveEntrySource& right) noexcept { return left.entryName == right.entryName; }),
                     outEntries.end());
    return outEntries.empty() ? HRESULT_FROM_WIN32(ERROR_NOT_FOUND) : S_OK;
}

[[nodiscard]] HRESULT WriteZipLocalHeader(
    HANDLE archiveFile, std::string_view entryNameUtf8, uint32_t crc32, uint32_t compressedSize, uint32_t uncompressedSize)
{
    if (entryNameUtf8.empty() || entryNameUtf8.size() > static_cast<size_t>((std::numeric_limits<uint16_t>::max)()))
    {
        return HRESULT_FROM_WIN32(ERROR_INVALID_NAME);
    }

    std::vector<std::byte> header;
    header.reserve(30u + entryNameUtf8.size());
    AppendLe32(header, kZipLocalFileHeaderSignature);
    AppendLe16(header, kZipVersionStored);
    AppendLe16(header, kZipGeneralPurposeUtf8);
    AppendLe16(header, kZipCompressionStored);
    AppendLe16(header, kZipDosTimeMidnight);
    AppendLe16(header, kZipDosDate1980Jan1);
    AppendLe32(header, crc32);
    AppendLe32(header, compressedSize);
    AppendLe32(header, uncompressedSize);
    AppendLe16(header, static_cast<uint16_t>(entryNameUtf8.size()));
    AppendLe16(header, 0u);
    for (const char ch : entryNameUtf8)
    {
        header.push_back(static_cast<std::byte>(static_cast<unsigned char>(ch)));
    }

    return WriteArchiveBytes(archiveFile, header);
}

[[nodiscard]] HRESULT PatchZipLocalHeaderSizes(
    HANDLE archiveFile, uint64_t localHeaderOffset, uint32_t crc32, uint32_t compressedSize, uint32_t uncompressedSize)
{
    std::vector<std::byte> patch;
    patch.reserve(12u);
    AppendLe32(patch, crc32);
    AppendLe32(patch, compressedSize);
    AppendLe32(patch, uncompressedSize);
    if (const HRESULT hr = SetArchiveFilePosition(archiveFile, localHeaderOffset + 14u); FAILED(hr))
    {
        return hr;
    }
    if (const HRESULT hr = WriteArchiveBytes(archiveFile, patch); FAILED(hr))
    {
        return hr;
    }
    return S_OK;
}

[[nodiscard]] HRESULT StreamFileIntoZip(HANDLE archiveFile, const std::filesystem::path& sourcePath, uint32_t& outCrc32, uint32_t& outSize) noexcept
{
    outCrc32 = 0u;
    outSize  = 0u;

#pragma warning(push)
#pragma warning(disable : 4625 4626) // WIL unique_hfile copy operations are intentionally deleted.
    wil::unique_hfile sourceFile(CreateFileW(sourcePath.c_str(),
                                             GENERIC_READ,
                                             FILE_SHARE_READ | FILE_SHARE_WRITE | FILE_SHARE_DELETE,
                                             nullptr,
                                             OPEN_EXISTING,
                                             FILE_ATTRIBUTE_NORMAL | FILE_FLAG_SEQUENTIAL_SCAN,
                                             nullptr));
#pragma warning(pop)
    if (! sourceFile)
    {
        return HResultFromLastError();
    }

    std::array<std::byte, 64u * 1024u> buffer{};
    uint32_t crc   = 0xFFFFFFFFu;
    uint64_t total = 0u;
    for (;;)
    {
        DWORD bytesRead = 0u;
        if (ReadFile(sourceFile.get(), buffer.data(), static_cast<DWORD>(buffer.size()), &bytesRead, nullptr) == FALSE)
        {
            return HResultFromLastError();
        }
        if (bytesRead == 0u)
        {
            break;
        }
        if (total + bytesRead > static_cast<uint64_t>((std::numeric_limits<uint32_t>::max)()))
        {
            return HRESULT_FROM_WIN32(ERROR_FILE_TOO_LARGE);
        }

        crc = ArchiveCrc32Update(crc, buffer.data(), bytesRead);
        if (const HRESULT hr = WriteArchiveBytes(archiveFile, buffer.data(), bytesRead); FAILED(hr))
        {
            return hr;
        }
        total += bytesRead;
    }

    outCrc32 = crc ^ 0xFFFFFFFFu;
    outSize  = static_cast<uint32_t>(total);
    return S_OK;
}

[[nodiscard]] HRESULT WriteZipCentralDirectoryEntry(HANDLE archiveFile, const ZipCentralEntry& entry)
{
    if (entry.entryNameUtf8.empty() || entry.entryNameUtf8.size() > static_cast<size_t>((std::numeric_limits<uint16_t>::max)()))
    {
        return HRESULT_FROM_WIN32(ERROR_INVALID_NAME);
    }

    std::vector<std::byte> header;
    header.reserve(46u + entry.entryNameUtf8.size());
    AppendLe32(header, kZipCentralDirectorySignature);
    AppendLe16(header, kZipVersionStored);
    AppendLe16(header, kZipVersionStored);
    AppendLe16(header, kZipGeneralPurposeUtf8);
    AppendLe16(header, kZipCompressionStored);
    AppendLe16(header, kZipDosTimeMidnight);
    AppendLe16(header, kZipDosDate1980Jan1);
    AppendLe32(header, entry.crc32);
    AppendLe32(header, entry.compressedSize);
    AppendLe32(header, entry.uncompressedSize);
    AppendLe16(header, static_cast<uint16_t>(entry.entryNameUtf8.size()));
    AppendLe16(header, 0u);
    AppendLe16(header, 0u);
    AppendLe16(header, 0u);
    AppendLe16(header, 0u);
    AppendLe32(header, entry.directory ? (static_cast<uint32_t>(kZipDirectoryExternalAttribute) << 16u) : 0u);
    AppendLe32(header, entry.localHeaderOffset);
    for (const char ch : entry.entryNameUtf8)
    {
        header.push_back(static_cast<std::byte>(static_cast<unsigned char>(ch)));
    }

    return WriteArchiveBytes(archiveFile, header);
}

[[nodiscard]] HRESULT WriteZipEndOfCentralDirectory(HANDLE archiveFile, uint16_t entryCount, uint32_t centralDirectorySize, uint32_t centralDirectoryOffset)
{
    std::vector<std::byte> eocd;
    eocd.reserve(kZipEndOfCentralDirMinSize);
    AppendLe32(eocd, kZipEndOfCentralDirSignature);
    AppendLe16(eocd, 0u);
    AppendLe16(eocd, 0u);
    AppendLe16(eocd, entryCount);
    AppendLe16(eocd, entryCount);
    AppendLe32(eocd, centralDirectorySize);
    AppendLe32(eocd, centralDirectoryOffset);
    AppendLe16(eocd, 0u);
    return WriteArchiveBytes(archiveFile, eocd);
}

[[nodiscard]] ArchiveOperationResult CreateStoredZipArchive(const std::vector<std::filesystem::path>& selectedPaths,
                                                            const std::filesystem::path& archivePath,
                                                            bool overwrite)
{
    ArchiveOperationResult result{};
    result.archivePath = archivePath;
    if (archivePath.empty())
    {
        result.hr = HRESULT_FROM_WIN32(ERROR_INVALID_NAME);
        return result;
    }

    if (! overwrite)
    {
        std::error_code existsEc;
        if (std::filesystem::exists(archivePath, existsEc))
        {
            result.hr = HRESULT_FROM_WIN32(ERROR_FILE_EXISTS);
            return result;
        }
        if (existsEc)
        {
            result.hr = HResultFromErrorCode(existsEc);
            return result;
        }
    }

    std::vector<ArchiveEntrySource> sources;
    result.hr = CollectArchiveSources(selectedPaths, sources);
    if (FAILED(result.hr))
    {
        return result;
    }
    if (sources.size() > static_cast<size_t>((std::numeric_limits<uint16_t>::max)()))
    {
        result.hr = HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
        return result;
    }

    const std::filesystem::path parent = archivePath.parent_path();
    if (! parent.empty())
    {
        std::error_code ec;
        std::filesystem::create_directories(parent, ec);
        if (ec)
        {
            result.hr = HResultFromErrorCode(ec);
            return result;
        }
    }

#pragma warning(push)
#pragma warning(disable : 4625 4626) // WIL unique_hfile copy operations are intentionally deleted.
    wil::unique_hfile archiveFile(CreateFileW(archivePath.c_str(),
                                              GENERIC_READ | GENERIC_WRITE,
                                              FILE_SHARE_READ,
                                              nullptr,
                                              overwrite ? CREATE_ALWAYS : CREATE_NEW,
                                              FILE_ATTRIBUTE_NORMAL | FILE_FLAG_SEQUENTIAL_SCAN,
                                              nullptr));
#pragma warning(pop)
    if (! archiveFile)
    {
        result.hr = HResultFromLastError();
        return result;
    }

    auto deleteIncompleteArchive = wil::scope_exit([&]
    {
        if (FAILED(result.hr))
        {
            archiveFile.reset();
            DeleteFileW(archivePath.c_str());
        }
    });

    std::vector<ZipCentralEntry> centralEntries;
    centralEntries.reserve(sources.size());

    for (const ArchiveEntrySource& source : sources)
    {
        std::string entryNameUtf8 = Utf8FromUtf16ForMakeFileList(source.entryName);
        if (entryNameUtf8.empty())
        {
            result.hr = HRESULT_FROM_WIN32(ERROR_NO_UNICODE_TRANSLATION);
            return result;
        }

        uint64_t localHeaderOffset = 0u;
        result.hr                  = GetArchiveFilePosition(archiveFile.get(), localHeaderOffset);
        if (FAILED(result.hr))
        {
            return result;
        }
        if (localHeaderOffset > static_cast<uint64_t>((std::numeric_limits<uint32_t>::max)()))
        {
            result.hr = HRESULT_FROM_WIN32(ERROR_FILE_TOO_LARGE);
            return result;
        }

        result.hr = WriteZipLocalHeader(archiveFile.get(), entryNameUtf8, 0u, 0u, 0u);
        if (FAILED(result.hr))
        {
            return result;
        }

        uint32_t crc32      = 0u;
        uint32_t storedSize = 0u;
        if (! source.directory)
        {
            result.hr = StreamFileIntoZip(archiveFile.get(), source.sourcePath, crc32, storedSize);
            if (FAILED(result.hr))
            {
                return result;
            }
        }

        uint64_t endPosition = 0u;
        result.hr            = GetArchiveFilePosition(archiveFile.get(), endPosition);
        if (FAILED(result.hr))
        {
            return result;
        }
        result.hr = PatchZipLocalHeaderSizes(archiveFile.get(), localHeaderOffset, crc32, storedSize, storedSize);
        if (FAILED(result.hr))
        {
            return result;
        }
        result.hr = SetArchiveFilePosition(archiveFile.get(), endPosition);
        if (FAILED(result.hr))
        {
            return result;
        }

        ZipCentralEntry central{};
        central.entryName         = source.entryName;
        central.entryNameUtf8     = std::move(entryNameUtf8);
        central.crc32             = crc32;
        central.compressedSize    = storedSize;
        central.uncompressedSize  = storedSize;
        central.localHeaderOffset = static_cast<uint32_t>(localHeaderOffset);
        central.directory         = source.directory;
        centralEntries.push_back(std::move(central));

        result.entries.push_back(source.entryName);
        ++result.entryCount;
        result.bytesProcessed += storedSize;
    }

    uint64_t centralDirectoryOffset = 0u;
    result.hr                       = GetArchiveFilePosition(archiveFile.get(), centralDirectoryOffset);
    if (FAILED(result.hr))
    {
        return result;
    }
    if (centralDirectoryOffset > static_cast<uint64_t>((std::numeric_limits<uint32_t>::max)()))
    {
        result.hr = HRESULT_FROM_WIN32(ERROR_FILE_TOO_LARGE);
        return result;
    }

    for (const ZipCentralEntry& entry : centralEntries)
    {
        result.hr = WriteZipCentralDirectoryEntry(archiveFile.get(), entry);
        if (FAILED(result.hr))
        {
            return result;
        }
    }

    uint64_t centralDirectoryEnd = 0u;
    result.hr                    = GetArchiveFilePosition(archiveFile.get(), centralDirectoryEnd);
    if (FAILED(result.hr))
    {
        return result;
    }
    const uint64_t centralDirectorySize = centralDirectoryEnd - centralDirectoryOffset;
    if (centralDirectorySize > static_cast<uint64_t>((std::numeric_limits<uint32_t>::max)()))
    {
        result.hr = HRESULT_FROM_WIN32(ERROR_FILE_TOO_LARGE);
        return result;
    }

    result.hr = WriteZipEndOfCentralDirectory(archiveFile.get(),
                                              static_cast<uint16_t>(centralEntries.size()),
                                              static_cast<uint32_t>(centralDirectorySize),
                                              static_cast<uint32_t>(centralDirectoryOffset));
    return result;
}

void SetArchivePropVariantBstr(PROPVARIANT* value, const std::wstring& text) noexcept
{
    if (! value)
    {
        return;
    }

    value->vt      = VT_BSTR;
    value->bstrVal = SysAllocStringLen(text.data(), static_cast<UINT>(text.size()));
    if (! value->bstrVal && ! text.empty())
    {
        value->vt = VT_EMPTY;
    }
}

void SetArchivePropVariantBool(PROPVARIANT* value, bool flag) noexcept
{
    if (! value)
    {
        return;
    }

    value->vt      = VT_BOOL;
    value->boolVal = flag ? VARIANT_TRUE : VARIANT_FALSE;
}

void SetArchivePropVariantUInt32(PROPVARIANT* value, UInt32 number) noexcept
{
    if (! value)
    {
        return;
    }

    value->vt    = VT_UI4;
    value->ulVal = number;
}

void SetArchivePropVariantUInt64(PROPVARIANT* value, UInt64 number) noexcept
{
    if (! value)
    {
        return;
    }

    value->vt             = VT_UI8;
    value->uhVal.QuadPart = number;
}

void SetArchivePropVariantFileTime(PROPVARIANT* value, const FILETIME& fileTime) noexcept
{
    if (! value)
    {
        return;
    }

    value->vt       = VT_FILETIME;
    value->filetime = fileTime;
}

[[nodiscard]] std::wstring SevenZipEntryPathProperty(const ArchiveEntrySource& source)
{
    std::wstring entryName = source.entryName;
    while (! entryName.empty() && entryName.back() == L'/')
    {
        entryName.pop_back();
    }
    std::replace(entryName.begin(), entryName.end(), L'/', L'\\');
    return entryName;
}

class SevenZipPackFileInStream final : public ISequentialInStream
{
public:
    static HRESULT Create(const std::filesystem::path& path, wil::com_ptr<ISequentialInStream>& out) noexcept
    {
        out.reset();
        if (path.empty())
        {
            return E_INVALIDARG;
        }

#pragma warning(push)
#pragma warning(disable : 4625 4626) // WIL unique_hfile copy operations are intentionally deleted.
        wil::unique_hfile file(CreateFileW(path.c_str(),
                                           GENERIC_READ,
                                           FILE_SHARE_READ | FILE_SHARE_WRITE | FILE_SHARE_DELETE,
                                           nullptr,
                                           OPEN_EXISTING,
                                           FILE_ATTRIBUTE_NORMAL | FILE_FLAG_SEQUENTIAL_SCAN,
                                           nullptr));
#pragma warning(pop)
        if (! file)
        {
            return HResultFromLastError();
        }

        auto* impl = new (std::nothrow) SevenZipPackFileInStream(std::move(file));
        if (! impl)
        {
            return E_OUTOFMEMORY;
        }

        out.attach(impl);
        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE QueryInterface(REFIID iid, void** outObject) noexcept override
    {
        if (! outObject)
        {
            return E_POINTER;
        }
        *outObject = nullptr;
        if (iid == IID_IUnknown || iid == IID_ISequentialInStream)
        {
            *outObject = static_cast<ISequentialInStream*>(this);
        }
        else
        {
            return E_NOINTERFACE;
        }

        AddRef();
        return S_OK;
    }

    ULONG STDMETHODCALLTYPE AddRef() noexcept override
    {
        return _refCount.fetch_add(1, std::memory_order_relaxed) + 1u;
    }

    ULONG STDMETHODCALLTYPE Release() noexcept override
    {
        const ULONG result = _refCount.fetch_sub(1, std::memory_order_acq_rel) - 1u;
        if (result == 0u)
        {
            delete this;
        }
        return result;
    }

    HRESULT STDMETHODCALLTYPE Read(void* data, UInt32 size, UInt32* processedSize) noexcept override
    {
        if (! processedSize)
        {
            return E_POINTER;
        }

        *processedSize = 0u;
        if (size == 0u)
        {
            return S_OK;
        }

        DWORD bytesRead = 0u;
        if (ReadFile(_file.get(), data, size, &bytesRead, nullptr) == FALSE)
        {
            return HResultFromLastError();
        }

        *processedSize = static_cast<UInt32>(bytesRead);
        return S_OK;
    }

private:
    explicit SevenZipPackFileInStream(wil::unique_hfile file) noexcept : _file(std::move(file))
    {
    }

    ~SevenZipPackFileInStream() = default;

    std::atomic_ulong _refCount{1u};
    wil::unique_hfile _file;
};

class SevenZipPackFileOutStream final : public IOutStream
{
public:
    static HRESULT Create(const std::filesystem::path& path, bool overwrite, wil::com_ptr<IOutStream>& out) noexcept
    {
        out.reset();
        if (path.empty())
        {
            return E_INVALIDARG;
        }

#pragma warning(push)
#pragma warning(disable : 4625 4626) // WIL unique_hfile copy operations are intentionally deleted.
        wil::unique_hfile file(CreateFileW(path.c_str(),
                                           GENERIC_READ | GENERIC_WRITE,
                                           FILE_SHARE_READ,
                                           nullptr,
                                           overwrite ? CREATE_ALWAYS : CREATE_NEW,
                                           FILE_ATTRIBUTE_NORMAL | FILE_FLAG_SEQUENTIAL_SCAN,
                                           nullptr));
#pragma warning(pop)
        if (! file)
        {
            return HResultFromLastError();
        }

        auto* impl = new (std::nothrow) SevenZipPackFileOutStream(std::move(file));
        if (! impl)
        {
            return E_OUTOFMEMORY;
        }

        out.attach(impl);
        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE QueryInterface(REFIID iid, void** outObject) noexcept override
    {
        if (! outObject)
        {
            return E_POINTER;
        }
        *outObject = nullptr;
        if (iid == IID_IUnknown || iid == IID_ISequentialOutStream)
        {
            *outObject = static_cast<ISequentialOutStream*>(this);
        }
        else if (iid == IID_IOutStream)
        {
            *outObject = static_cast<IOutStream*>(this);
        }
        else
        {
            return E_NOINTERFACE;
        }

        AddRef();
        return S_OK;
    }

    ULONG STDMETHODCALLTYPE AddRef() noexcept override
    {
        return _refCount.fetch_add(1, std::memory_order_relaxed) + 1u;
    }

    ULONG STDMETHODCALLTYPE Release() noexcept override
    {
        const ULONG result = _refCount.fetch_sub(1, std::memory_order_acq_rel) - 1u;
        if (result == 0u)
        {
            delete this;
        }
        return result;
    }

    HRESULT STDMETHODCALLTYPE Write(const void* data, UInt32 size, UInt32* processedSize) noexcept override
    {
        if (! processedSize)
        {
            return E_POINTER;
        }

        *processedSize = 0u;
        if (size == 0u)
        {
            return S_OK;
        }

        DWORD bytesWritten = 0u;
        if (WriteFile(_file.get(), data, size, &bytesWritten, nullptr) == FALSE)
        {
            return HResultFromLastError();
        }

        *processedSize = static_cast<UInt32>(bytesWritten);
        return bytesWritten == 0u ? HRESULT_FROM_WIN32(ERROR_WRITE_FAULT) : S_OK;
    }

    HRESULT STDMETHODCALLTYPE Seek(Int64 offset, UInt32 seekOrigin, UInt64* newPosition) noexcept override
    {
        DWORD method = FILE_BEGIN;
        switch (seekOrigin)
        {
            case 0u: method = FILE_BEGIN; break;
            case 1u: method = FILE_CURRENT; break;
            case 2u: method = FILE_END; break;
            default: return STG_E_INVALIDFUNCTION;
        }

        LARGE_INTEGER distance{};
        distance.QuadPart = offset;
        LARGE_INTEGER position{};
        if (SetFilePointerEx(_file.get(), distance, &position, method) == FALSE)
        {
            return HResultFromLastError();
        }
        if (position.QuadPart < 0)
        {
            return HRESULT_FROM_WIN32(ERROR_NEGATIVE_SEEK);
        }
        if (newPosition)
        {
            *newPosition = static_cast<UInt64>(position.QuadPart);
        }
        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE SetSize(UInt64 newSize) noexcept override
    {
        if (newSize > static_cast<UInt64>((std::numeric_limits<LONGLONG>::max)()))
        {
            return HRESULT_FROM_WIN32(ERROR_FILE_TOO_LARGE);
        }

        LARGE_INTEGER current{};
        LARGE_INTEGER zero{};
        if (SetFilePointerEx(_file.get(), zero, &current, FILE_CURRENT) == FALSE)
        {
            return HResultFromLastError();
        }

        LARGE_INTEGER size{};
        size.QuadPart = static_cast<LONGLONG>(newSize);
        if (SetFilePointerEx(_file.get(), size, nullptr, FILE_BEGIN) == FALSE)
        {
            return HResultFromLastError();
        }
        if (SetEndOfFile(_file.get()) == FALSE)
        {
            return HResultFromLastError();
        }
        return SetFilePointerEx(_file.get(), current, nullptr, FILE_BEGIN) == FALSE ? HResultFromLastError() : S_OK;
    }

private:
    explicit SevenZipPackFileOutStream(wil::unique_hfile file) noexcept : _file(std::move(file))
    {
    }

    ~SevenZipPackFileOutStream() = default;

    std::atomic_ulong _refCount{1u};
    wil::unique_hfile _file;
};

class SevenZipUpdateCallback final : public IArchiveUpdateCallback
{
public:
    explicit SevenZipUpdateCallback(const std::vector<ArchiveEntrySource>& sources) noexcept : _sources(sources)
    {
    }

    [[nodiscard]] HRESULT ResultHr() const noexcept
    {
        return _resultHr;
    }

    HRESULT STDMETHODCALLTYPE QueryInterface(REFIID iid, void** outObject) noexcept override
    {
        if (! outObject)
        {
            return E_POINTER;
        }
        *outObject = nullptr;
        if (iid == IID_IUnknown || iid == IID_IProgress || iid == IID_IArchiveUpdateCallback)
        {
            *outObject = static_cast<IArchiveUpdateCallback*>(this);
        }
        else
        {
            return E_NOINTERFACE;
        }

        AddRef();
        return S_OK;
    }

    ULONG STDMETHODCALLTYPE AddRef() noexcept override
    {
        return _refCount.fetch_add(1, std::memory_order_relaxed) + 1u;
    }

    ULONG STDMETHODCALLTYPE Release() noexcept override
    {
        const ULONG result = _refCount.fetch_sub(1, std::memory_order_acq_rel) - 1u;
        if (result == 0u)
        {
            delete this;
        }
        return result;
    }

    HRESULT STDMETHODCALLTYPE SetTotal(UInt64) noexcept override
    {
        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE SetCompleted(const UInt64*) noexcept override
    {
        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE GetUpdateItemInfo(UInt32 index, Int32* newData, Int32* newProps, UInt32* indexInArchive) noexcept override
    {
        if (index >= _sources.size())
        {
            return E_INVALIDARG;
        }
        if (newData)
        {
            *newData = 1;
        }
        if (newProps)
        {
            *newProps = 1;
        }
        if (indexInArchive)
        {
            *indexInArchive = static_cast<UInt32>(-1);
        }
        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE GetProperty(UInt32 index, PROPID propID, PROPVARIANT* value) noexcept override
    {
        if (! value)
        {
            return E_POINTER;
        }
        PropVariantInit(value);
        if (index >= _sources.size())
        {
            return E_INVALIDARG;
        }

        const ArchiveEntrySource& source = _sources[index];
        switch (propID)
        {
            case kpidPath: SetArchivePropVariantBstr(value, SevenZipEntryPathProperty(source)); break;
            case kpidIsDir: SetArchivePropVariantBool(value, source.directory); break;
            case kpidSize:
                if (! source.directory)
                {
                    SetArchivePropVariantUInt64(value, source.fileSize);
                }
                break;
            case kpidAttrib:
            {
                DWORD attributes = GetFileAttributesW(source.sourcePath.c_str());
                if (attributes == INVALID_FILE_ATTRIBUTES)
                {
                    attributes = source.directory ? FILE_ATTRIBUTE_DIRECTORY : FILE_ATTRIBUTE_ARCHIVE;
                }
                SetArchivePropVariantUInt32(value, attributes);
                break;
            }
            case kpidMTime:
            {
                WIN32_FILE_ATTRIBUTE_DATA data{};
                if (GetFileAttributesExW(source.sourcePath.c_str(), GetFileExInfoStandard, &data) != FALSE)
                {
                    SetArchivePropVariantFileTime(value, data.ftLastWriteTime);
                }
                break;
            }
            default: break;
        }

        return value->vt == VT_EMPTY && propID == kpidPath ? E_OUTOFMEMORY : S_OK;
    }

    HRESULT STDMETHODCALLTYPE GetStream(UInt32 index, ISequentialInStream** inStream) noexcept override
    {
        if (! inStream)
        {
            return E_POINTER;
        }
        *inStream = nullptr;
        if (index >= _sources.size())
        {
            return E_INVALIDARG;
        }

        const ArchiveEntrySource& source = _sources[index];
        if (source.directory)
        {
            return S_OK;
        }

        wil::com_ptr<ISequentialInStream> stream;
        const HRESULT hr = SevenZipPackFileInStream::Create(source.sourcePath, stream);
        if (FAILED(hr))
        {
            _resultHr = hr;
            return S_FALSE;
        }

        *inStream = stream.detach();
        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE SetOperationResult(Int32 operationResult) noexcept override
    {
        if (operationResult != NArchive::NUpdate::NOperationResult::kOK && SUCCEEDED(_resultHr))
        {
            _resultHr = E_FAIL;
        }
        return S_OK;
    }

private:
    std::atomic_ulong _refCount{1u};
    const std::vector<ArchiveEntrySource>& _sources;
    HRESULT _resultHr = S_OK;
};

[[nodiscard]] ArchiveOperationResult CreateSevenZipArchive(const std::vector<std::filesystem::path>& selectedPaths,
                                                           const std::filesystem::path& archivePath,
                                                           bool overwrite,
                                                           const ArchivePackerDefinition& packer)
{
    ArchiveOperationResult result{};
    result.archivePath = archivePath;
    if (archivePath.empty() || ! packer.hasClassId)
    {
        result.hr = HRESULT_FROM_WIN32(ERROR_INVALID_NAME);
        return result;
    }

    if (! overwrite)
    {
        std::error_code existsEc;
        if (std::filesystem::exists(archivePath, existsEc))
        {
            result.hr = HRESULT_FROM_WIN32(ERROR_FILE_EXISTS);
            return result;
        }
        if (existsEc)
        {
            result.hr = HResultFromErrorCode(existsEc);
            return result;
        }
    }

    std::vector<ArchiveEntrySource> sources;
    result.hr = CollectArchiveSources(selectedPaths, sources);
    if (FAILED(result.hr))
    {
        return result;
    }
    if (sources.size() > static_cast<size_t>((std::numeric_limits<UInt32>::max)()))
    {
        result.hr = HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
        return result;
    }

    const std::filesystem::path parent = archivePath.parent_path();
    if (! parent.empty())
    {
        std::error_code ec;
        std::filesystem::create_directories(parent, ec);
        if (ec)
        {
            result.hr = HResultFromErrorCode(ec);
            return result;
        }
    }

    SevenZipArchiveExports& api = GetSevenZipArchiveExports();
    result.hr                   = LoadSevenZipArchiveExports(api);
    if (FAILED(result.hr) || ! api.createObject)
    {
        return result;
    }

    wil::com_ptr<IOutArchive> archive;
    result.hr = api.createObject(&packer.classId, &IID_IOutArchive, archive.put_void());
    if (FAILED(result.hr) || ! archive)
    {
        if (SUCCEEDED(result.hr))
        {
            result.hr = E_NOINTERFACE;
        }
        return result;
    }

    wil::com_ptr<IOutStream> outputStream;
    result.hr = SevenZipPackFileOutStream::Create(archivePath, overwrite, outputStream);
    if (FAILED(result.hr))
    {
        return result;
    }

    auto* callbackImpl = new (std::nothrow) SevenZipUpdateCallback(sources);
    if (! callbackImpl)
    {
        result.hr = E_OUTOFMEMORY;
        return result;
    }

    wil::com_ptr<IArchiveUpdateCallback> callback;
    callback.attach(callbackImpl);
    result.hr = archive->UpdateItems(outputStream.get(), static_cast<UInt32>(sources.size()), callback.get());
    if (SUCCEEDED(result.hr) && FAILED(callbackImpl->ResultHr()))
    {
        result.hr = callbackImpl->ResultHr();
    }

    outputStream.reset();
    if (FAILED(result.hr))
    {
        DeleteFileW(archivePath.c_str());
        return result;
    }

    for (const ArchiveEntrySource& source : sources)
    {
        result.entries.push_back(source.entryName);
        ++result.entryCount;
        if (! source.directory)
        {
            result.bytesProcessed += source.fileSize;
        }
    }

    return result;
}

[[nodiscard]] HRESULT FindZipEndOfCentralDirectory(HANDLE archiveFile, uint64_t fileSize, ZipEndOfCentralDirectory& outEocd)
{
    outEocd = {};
    if (fileSize < kZipEndOfCentralDirMinSize)
    {
        return HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
    }

    const size_t tailSize     = static_cast<size_t>((std::min)(fileSize, static_cast<uint64_t>(kZipMaxTailSearchSize)));
    const uint64_t tailOffset = fileSize - tailSize;
    std::vector<std::byte> tail;
    if (const HRESULT hr = ReadArchiveBytesAt(archiveFile, tailOffset, tailSize, tail); FAILED(hr))
    {
        return hr;
    }

    bool sawUnsupportedCandidate = false;
    for (size_t offset = tail.size() - kZipEndOfCentralDirMinSize + 1u; offset > 0u; --offset)
    {
        const size_t index = offset - 1u;
        uint32_t signature = 0u;
        if (! TryReadLe32(tail, index, signature) || signature != kZipEndOfCentralDirSignature)
        {
            continue;
        }

        uint16_t diskNumber      = 0u;
        uint16_t centralDisk     = 0u;
        uint16_t diskEntryCount  = 0u;
        uint16_t totalEntryCount = 0u;
        uint16_t commentLength   = 0u;
        uint32_t centralSize     = 0u;
        uint32_t centralOffset   = 0u;
        if (! TryReadLe16(tail, index + 4u, diskNumber) || ! TryReadLe16(tail, index + 6u, centralDisk) || ! TryReadLe16(tail, index + 8u, diskEntryCount) ||
            ! TryReadLe16(tail, index + 10u, totalEntryCount) || ! TryReadLe32(tail, index + 12u, centralSize) ||
            ! TryReadLe32(tail, index + 16u, centralOffset) || ! TryReadLe16(tail, index + 20u, commentLength))
        {
            continue;
        }

        if (diskNumber != 0u || centralDisk != 0u || diskEntryCount != totalEntryCount)
        {
            sawUnsupportedCandidate = true;
            continue;
        }
        if (index + kZipEndOfCentralDirMinSize + commentLength > tail.size())
        {
            continue;
        }
        if (static_cast<uint64_t>(centralOffset) + centralSize > fileSize)
        {
            continue;
        }

        outEocd.entryCount             = totalEntryCount;
        outEocd.centralDirectorySize   = centralSize;
        outEocd.centralDirectoryOffset = centralOffset;
        return S_OK;
    }

    return sawUnsupportedCandidate ? HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED) : HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
}

[[nodiscard]] std::wstring Utf16FromArchiveCodePage(std::string_view text, UINT codePage, DWORD flags) noexcept
{
    if (text.empty() || text.size() > static_cast<size_t>((std::numeric_limits<int>::max)()))
    {
        return {};
    }

    const int inputBytes = static_cast<int>(text.size());
    const int required   = MultiByteToWideChar(codePage, flags, text.data(), inputBytes, nullptr, 0);
    if (required <= 0)
    {
        return {};
    }

    std::wstring result(static_cast<size_t>(required), L'\0');
    const int written = MultiByteToWideChar(codePage, flags, text.data(), inputBytes, result.data(), required);
    if (written != required)
    {
        return {};
    }

    return result;
}

[[nodiscard]] std::wstring DecodeZipEntryName(std::string_view rawName, uint16_t flags) noexcept
{
    constexpr uint16_t kZipGeneralPurposeUtf8Flag = 0x0800u;
    constexpr UINT kZipLegacyCodePageCp437        = 437u;
    if ((flags & kZipGeneralPurposeUtf8Flag) != 0u)
    {
        return Utf16FromArchiveCodePage(rawName, CP_UTF8, MB_ERR_INVALID_CHARS);
    }

    return Utf16FromArchiveCodePage(rawName, kZipLegacyCodePageCp437, 0u);
}

[[nodiscard]] HRESULT ParseZipCentralDirectory(HANDLE archiveFile, const ZipEndOfCentralDirectory& eocd, std::vector<ZipParsedEntry>& outEntries)
{
    outEntries.clear();
    std::vector<std::byte> centralBytes;
    if (const HRESULT hr = ReadArchiveBytesAt(archiveFile, eocd.centralDirectoryOffset, eocd.centralDirectorySize, centralBytes); FAILED(hr))
    {
        return hr;
    }

    size_t offset                  = 0u;
    uint64_t totalUncompressedSize = 0u;
    outEntries.reserve(eocd.entryCount);
    for (uint16_t entryIndex = 0u; entryIndex < eocd.entryCount; ++entryIndex)
    {
        uint32_t signature         = 0u;
        uint16_t flags             = 0u;
        uint16_t compressionMethod = 0u;
        uint32_t crc32             = 0u;
        uint32_t compressedSize    = 0u;
        uint32_t uncompressedSize  = 0u;
        uint16_t nameLength        = 0u;
        uint16_t extraLength       = 0u;
        uint16_t commentLength     = 0u;
        uint16_t diskStart         = 0u;
        uint32_t localHeaderOffset = 0u;
        if (! TryReadLe32(centralBytes, offset, signature) || ! TryReadLe16(centralBytes, offset + 8u, flags) ||
            ! TryReadLe16(centralBytes, offset + 10u, compressionMethod) || ! TryReadLe32(centralBytes, offset + 16u, crc32) ||
            ! TryReadLe32(centralBytes, offset + 20u, compressedSize) || ! TryReadLe32(centralBytes, offset + 24u, uncompressedSize) ||
            ! TryReadLe16(centralBytes, offset + 28u, nameLength) || ! TryReadLe16(centralBytes, offset + 30u, extraLength) ||
            ! TryReadLe16(centralBytes, offset + 32u, commentLength) || ! TryReadLe16(centralBytes, offset + 34u, diskStart) ||
            ! TryReadLe32(centralBytes, offset + 42u, localHeaderOffset))
        {
            return HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
        }
        if (signature != kZipCentralDirectorySignature)
        {
            return HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
        }
        if ((flags & 0x0001u) != 0u || diskStart != 0u)
        {
            return HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
        }
        if (compressionMethod != kZipCompressionStored)
        {
            return HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
        }
        if (compressedSize != uncompressedSize)
        {
            return HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
        }

        const size_t variableOffset = offset + 46u;
        const size_t nextOffset     = variableOffset + nameLength + extraLength + commentLength;
        if (variableOffset > centralBytes.size() || nextOffset > centralBytes.size())
        {
            return HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
        }

        std::string rawName;
        rawName.resize(nameLength);
        for (size_t i = 0u; i < nameLength; ++i)
        {
            rawName[i] = static_cast<char>(std::to_integer<unsigned char>(centralBytes[variableOffset + i]));
        }

        std::wstring entryName = DecodeZipEntryName(rawName, flags);
        if (entryName.empty())
        {
            return HRESULT_FROM_WIN32(ERROR_NO_UNICODE_TRANSLATION);
        }

        bool directory = false;
        std::filesystem::path relativePath;
        if (! IsArchiveEntryNameSafe(entryName, directory, relativePath))
        {
            return HRESULT_FROM_WIN32(ERROR_INVALID_NAME);
        }
        if (! directory)
        {
            if (const HRESULT sizeHr = AccumulateArchiveExtractSize(totalUncompressedSize, uncompressedSize); FAILED(sizeHr))
            {
                return sizeHr;
            }
        }

        ZipParsedEntry parsed{};
        parsed.entryName         = std::move(entryName);
        parsed.relativePath      = std::move(relativePath);
        parsed.crc32             = crc32;
        parsed.compressedSize    = compressedSize;
        parsed.uncompressedSize  = uncompressedSize;
        parsed.localHeaderOffset = localHeaderOffset;
        parsed.directory         = directory;
        outEntries.push_back(std::move(parsed));
        offset = nextOffset;
    }

    return offset == centralBytes.size() ? S_OK : HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
}

[[nodiscard]] HRESULT ResolveZipEntryDataOffset(HANDLE archiveFile, const ZipParsedEntry& entry, uint64_t fileSize, uint64_t& outDataOffset)
{
    outDataOffset = 0u;
    std::array<std::byte, 30u> header{};
    if (const HRESULT hr = ReadArchiveBytesAt(archiveFile, entry.localHeaderOffset, std::span<std::byte>(header.data(), header.size())); FAILED(hr))
    {
        return hr;
    }

    uint32_t signature   = 0u;
    uint16_t nameLength  = 0u;
    uint16_t extraLength = 0u;
    if (! TryReadLe32(header, 0u, signature) || ! TryReadLe16(header, 26u, nameLength) || ! TryReadLe16(header, 28u, extraLength))
    {
        return HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
    }
    if (signature != kZipLocalFileHeaderSignature)
    {
        return HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
    }

    outDataOffset = static_cast<uint64_t>(entry.localHeaderOffset) + header.size() + nameLength + extraLength;
    if (outDataOffset > fileSize || outDataOffset + entry.compressedSize > fileSize)
    {
        return HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
    }
    return S_OK;
}

[[nodiscard]] HRESULT ExtractStoredZipFileEntry(HANDLE archiveFile,
                                                const ZipParsedEntry& entry,
                                                uint64_t dataOffset,
                                                const std::filesystem::path& targetPath,
                                                ArchiveExistingTargetPolicy conflictPolicy,
                                                uint32_t& outCrc32) noexcept
{
    outCrc32 = 0u;
    Common::Files::LocalFileTransaction transaction;
    HRESULT transactionHr = Common::Files::LocalFileTransaction::Create(targetPath, LocalFilePolicyForArchive(conflictPolicy), true, transaction);
    if (FAILED(transactionHr))
    {
        return conflictPolicy == ArchiveExistingTargetPolicy::Skip && IsArchiveTargetExistsFailure(transactionHr) ? S_FALSE : transactionHr;
    }

    std::array<std::byte, 64u * 1024u> buffer{};
    uint32_t crc        = 0xFFFFFFFFu;
    uint64_t remaining  = entry.compressedSize;
    uint64_t readOffset = dataOffset;
    while (remaining > 0u)
    {
        const size_t chunk = static_cast<size_t>((std::min)(remaining, static_cast<uint64_t>(buffer.size())));
        if (const HRESULT hr = ReadArchiveBytesAt(archiveFile, readOffset, std::span<std::byte>(buffer.data(), chunk)); FAILED(hr))
        {
            return hr;
        }
        crc = ArchiveCrc32Update(crc, buffer.data(), chunk);
        if (const HRESULT hr = transaction.Write(std::span<const std::byte>(buffer.data(), chunk)); FAILED(hr))
        {
            return hr;
        }
        readOffset += chunk;
        remaining -= chunk;
    }

    outCrc32 = crc ^ 0xFFFFFFFFu;
    if (outCrc32 != entry.crc32)
    {
        return HRESULT_FROM_WIN32(ERROR_CRC);
    }

    transactionHr = transaction.Commit(entry.uncompressedSize);
    if (FAILED(transactionHr))
    {
        return conflictPolicy == ArchiveExistingTargetPolicy::Skip && IsArchiveTargetExistsFailure(transactionHr) ? S_FALSE : transactionHr;
    }
    return S_OK;
}

[[nodiscard]] bool ArchiveUnpackMaskMatchesAll(std::wstring_view maskText) noexcept
{
    const std::wstring trimmed = StringUtils::TrimWhitespaceCopy(maskText);
    return trimmed.empty() || trimmed == L"*.*" || trimmed == L"*";
}

[[nodiscard]] bool ShouldExtractArchiveEntryForMask(std::wstring_view entryNameView,
                                                    const std::filesystem::path& relativePath,
                                                    const MaskSyntax::WildcardMask& mask,
                                                    bool maskMatchesAll)
{
    if (maskMatchesAll)
    {
        return true;
    }

    std::wstring entryName(entryNameView);
    while (! entryName.empty() && entryName.back() == L'/')
    {
        entryName.pop_back();
    }

    const std::wstring fileName = relativePath.filename().wstring();
    return (! fileName.empty() && MaskSyntax::MatchesWildcardMask(fileName, mask)) || (! entryName.empty() && MaskSyntax::MatchesWildcardMask(entryName, mask));
}

[[nodiscard]] bool ShouldExtractZipEntryForMask(const ZipParsedEntry& entry, const MaskSyntax::WildcardMask& mask, bool maskMatchesAll)
{
    return ShouldExtractArchiveEntryForMask(entry.entryName, entry.relativePath, mask, maskMatchesAll);
}

struct SevenZipExtractEntry final
{
    UInt32 archiveIndex = 0u;
    std::wstring entryName;
    std::filesystem::path relativePath;
    std::filesystem::path targetPath;
    uint64_t sizeBytes = 0u;
    bool directory     = false;
};

class SevenZipExtractFileInStream final : public IInStream, public IStreamGetSize
{
public:
    static HRESULT Create(const std::filesystem::path& path, wil::com_ptr<IInStream>& out) noexcept
    {
        out.reset();
        if (path.empty())
        {
            return E_INVALIDARG;
        }

#pragma warning(push)
#pragma warning(disable : 4625 4626) // WIL unique_hfile copy operations are intentionally deleted.
        wil::unique_hfile file(CreateFileW(path.c_str(),
                                           GENERIC_READ,
                                           FILE_SHARE_READ | FILE_SHARE_WRITE | FILE_SHARE_DELETE,
                                           nullptr,
                                           OPEN_EXISTING,
                                           FILE_ATTRIBUTE_NORMAL | FILE_FLAG_SEQUENTIAL_SCAN,
                                           nullptr));
#pragma warning(pop)
        if (! file)
        {
            return HResultFromLastError();
        }

        LARGE_INTEGER fileSize{};
        if (GetFileSizeEx(file.get(), &fileSize) == FALSE)
        {
            return HResultFromLastError();
        }
        if (fileSize.QuadPart < 0)
        {
            return HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
        }

        auto* impl = new (std::nothrow) SevenZipExtractFileInStream(std::move(file), static_cast<uint64_t>(fileSize.QuadPart));
        if (! impl)
        {
            return E_OUTOFMEMORY;
        }

        out.attach(impl);
        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE QueryInterface(REFIID iid, void** outObject) noexcept override
    {
        if (! outObject)
        {
            return E_POINTER;
        }
        *outObject = nullptr;
        if (iid == IID_IUnknown || iid == IID_IInStream)
        {
            *outObject = static_cast<IInStream*>(this);
        }
        else if (iid == IID_ISequentialInStream)
        {
            *outObject = static_cast<ISequentialInStream*>(this);
        }
        else if (iid == IID_IStreamGetSize)
        {
            *outObject = static_cast<IStreamGetSize*>(this);
        }
        else
        {
            return E_NOINTERFACE;
        }

        AddRef();
        return S_OK;
    }

    ULONG STDMETHODCALLTYPE AddRef() noexcept override
    {
        return _refCount.fetch_add(1u, std::memory_order_relaxed) + 1u;
    }

    ULONG STDMETHODCALLTYPE Release() noexcept override
    {
        const ULONG result = _refCount.fetch_sub(1u, std::memory_order_acq_rel) - 1u;
        if (result == 0u)
        {
            delete this;
        }
        return result;
    }

    HRESULT STDMETHODCALLTYPE Read(void* data, UInt32 size, UInt32* processedSize) noexcept override
    {
        if (! processedSize)
        {
            return E_POINTER;
        }

        *processedSize = 0u;
        if (size == 0u)
        {
            return S_OK;
        }
        if (! data)
        {
            return E_POINTER;
        }

        DWORD bytesRead = 0u;
        if (ReadFile(_file.get(), data, size, &bytesRead, nullptr) == FALSE)
        {
            return HResultFromLastError();
        }

        *processedSize = static_cast<UInt32>(bytesRead);
        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE Seek(Int64 offset, UInt32 seekOrigin, UInt64* newPosition) noexcept override
    {
        DWORD method = FILE_BEGIN;
        switch (seekOrigin)
        {
            case 0u: method = FILE_BEGIN; break;
            case 1u: method = FILE_CURRENT; break;
            case 2u: method = FILE_END; break;
            default: return STG_E_INVALIDFUNCTION;
        }

        LARGE_INTEGER distance{};
        distance.QuadPart = offset;
        LARGE_INTEGER position{};
        if (SetFilePointerEx(_file.get(), distance, &position, method) == FALSE)
        {
            return HResultFromLastError();
        }
        if (position.QuadPart < 0)
        {
            return HRESULT_FROM_WIN32(ERROR_NEGATIVE_SEEK);
        }
        if (newPosition)
        {
            *newPosition = static_cast<UInt64>(position.QuadPart);
        }
        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE GetSize(UInt64* size) noexcept override
    {
        if (! size)
        {
            return E_POINTER;
        }
        *size = static_cast<UInt64>(_sizeBytes);
        return S_OK;
    }

private:
    SevenZipExtractFileInStream(wil::unique_hfile file, uint64_t sizeBytes) noexcept : _file(std::move(file)), _sizeBytes(sizeBytes)
    {
    }

    ~SevenZipExtractFileInStream() = default;

    std::atomic_ulong _refCount{1u};
    wil::unique_hfile _file;
    uint64_t _sizeBytes = 0u;
};

class SevenZipExtractFileOutStream final : public ISequentialOutStream
{
public:
    static HRESULT Create(const std::filesystem::path& targetPath,
                          ArchiveExistingTargetPolicy conflictPolicy,
                          wil::com_ptr<SevenZipExtractFileOutStream>& out) noexcept
    {
        out.reset();
        if (targetPath.empty())
        {
            return HRESULT_FROM_WIN32(ERROR_INVALID_NAME);
        }

        Common::Files::LocalFileTransaction transaction;
        const HRESULT createHr = Common::Files::LocalFileTransaction::Create(targetPath, LocalFilePolicyForArchive(conflictPolicy), true, transaction);
        if (FAILED(createHr))
        {
            return createHr;
        }

        auto* impl = new (std::nothrow) SevenZipExtractFileOutStream(std::move(transaction));
        if (! impl)
        {
            return E_OUTOFMEMORY;
        }

        out.attach(impl);
        return S_OK;
    }

    HRESULT Commit() noexcept
    {
        return _transaction.Commit(_bytesWritten);
    }

    [[nodiscard]] uint64_t BytesWritten() const noexcept
    {
        return _bytesWritten;
    }

    HRESULT STDMETHODCALLTYPE QueryInterface(REFIID iid, void** outObject) noexcept override
    {
        if (! outObject)
        {
            return E_POINTER;
        }
        *outObject = nullptr;
        if (iid == IID_IUnknown || iid == IID_ISequentialOutStream)
        {
            *outObject = static_cast<ISequentialOutStream*>(this);
        }
        else
        {
            return E_NOINTERFACE;
        }

        AddRef();
        return S_OK;
    }

    ULONG STDMETHODCALLTYPE AddRef() noexcept override
    {
        return _refCount.fetch_add(1u, std::memory_order_relaxed) + 1u;
    }

    ULONG STDMETHODCALLTYPE Release() noexcept override
    {
        const ULONG result = _refCount.fetch_sub(1u, std::memory_order_acq_rel) - 1u;
        if (result == 0u)
        {
            delete this;
        }
        return result;
    }

    HRESULT STDMETHODCALLTYPE Write(const void* data, UInt32 size, UInt32* processedSize) noexcept override
    {
        if (! processedSize)
        {
            return E_POINTER;
        }

        *processedSize = 0u;
        if (size == 0u)
        {
            return S_OK;
        }
        if (! data)
        {
            return E_POINTER;
        }
        if (static_cast<uint64_t>(size) > kMaxArchiveExtractEntryBytes - _bytesWritten)
        {
            return HRESULT_FROM_WIN32(ERROR_FILE_TOO_LARGE);
        }

        const HRESULT writeHr = _transaction.Write(data, size);
        if (FAILED(writeHr))
        {
            return writeHr;
        }

        *processedSize = size;
        _bytesWritten += static_cast<uint64_t>(size);
        return S_OK;
    }

private:
    explicit SevenZipExtractFileOutStream(Common::Files::LocalFileTransaction transaction) noexcept : _transaction(std::move(transaction))
    {
    }

    ~SevenZipExtractFileOutStream() noexcept = default;

    std::atomic_ulong _refCount{1u};
    Common::Files::LocalFileTransaction _transaction;
    uint64_t _bytesWritten = 0u;
};

[[nodiscard]] HRESULT SevenZipExtractOperationResultToHr(Int32 operationResult) noexcept
{
    switch (operationResult)
    {
        case NArchive::NExtract::NOperationResult::kOK: return S_OK;
        case NArchive::NExtract::NOperationResult::kUnsupportedMethod: return HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
        case NArchive::NExtract::NOperationResult::kCRCError: return HRESULT_FROM_WIN32(ERROR_CRC);
        case NArchive::NExtract::NOperationResult::kWrongPassword: return HRESULT_FROM_WIN32(ERROR_INVALID_PASSWORD);
        case NArchive::NExtract::NOperationResult::kUnavailable: return HRESULT_FROM_WIN32(ERROR_FILE_NOT_FOUND);
        case NArchive::NExtract::NOperationResult::kUnexpectedEnd: return HRESULT_FROM_WIN32(ERROR_HANDLE_EOF);
        case NArchive::NExtract::NOperationResult::kDataError:
        case NArchive::NExtract::NOperationResult::kDataAfterEnd:
        case NArchive::NExtract::NOperationResult::kIsNotArc:
        case NArchive::NExtract::NOperationResult::kHeadersError: return HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
        default: return E_FAIL;
    }
}

class SevenZipExtractCallback final : public IArchiveExtractCallback, public ICryptoGetTextPassword, public ICryptoGetTextPassword2
{
public:
    SevenZipExtractCallback(const std::vector<SevenZipExtractEntry>& entries,
                            ArchiveExistingTargetPolicy conflictPolicy,
                            ArchiveOperationResult& result) noexcept
        : _entries(entries),
          _conflictPolicy(conflictPolicy),
          _result(result)
    {
    }

    [[nodiscard]] HRESULT ResultHr() const noexcept
    {
        return _resultHr;
    }

    HRESULT STDMETHODCALLTYPE QueryInterface(REFIID iid, void** outObject) noexcept override
    {
        if (! outObject)
        {
            return E_POINTER;
        }
        *outObject = nullptr;
        if (iid == IID_IUnknown || iid == IID_IProgress)
        {
            *outObject = static_cast<IProgress*>(this);
        }
        else if (iid == IID_IArchiveExtractCallback)
        {
            *outObject = static_cast<IArchiveExtractCallback*>(this);
        }
        else if (iid == IID_ICryptoGetTextPassword)
        {
            *outObject = static_cast<ICryptoGetTextPassword*>(this);
        }
        else if (iid == IID_ICryptoGetTextPassword2)
        {
            *outObject = static_cast<ICryptoGetTextPassword2*>(this);
        }
        else
        {
            return E_NOINTERFACE;
        }

        AddRef();
        return S_OK;
    }

    ULONG STDMETHODCALLTYPE AddRef() noexcept override
    {
        return _refCount.fetch_add(1u, std::memory_order_relaxed) + 1u;
    }

    ULONG STDMETHODCALLTYPE Release() noexcept override
    {
        const ULONG result = _refCount.fetch_sub(1u, std::memory_order_acq_rel) - 1u;
        if (result == 0u)
        {
            delete this;
        }
        return result;
    }

    HRESULT STDMETHODCALLTYPE SetTotal(UInt64) noexcept override
    {
        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE SetCompleted(const UInt64*) noexcept override
    {
        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE GetStream(UInt32 index, ISequentialOutStream** outStream, Int32 askExtractMode) noexcept override
    {
        if (! outStream)
        {
            return E_POINTER;
        }

        *outStream   = nullptr;
        _activeEntry = nullptr;
        _activeFileStream.reset();
        if (askExtractMode != NArchive::NExtract::NAskMode::kExtract)
        {
            return S_OK;
        }

        const SevenZipExtractEntry* entry = FindEntry(index);
        if (! entry)
        {
            return S_OK;
        }

        _activeEntry = entry;
        if (entry->directory)
        {
            std::error_code ec;
            std::filesystem::create_directories(entry->targetPath, ec);
            if (ec)
            {
                _resultHr = HResultFromErrorCode(ec);
                return _resultHr;
            }
            return S_OK;
        }

        wil::com_ptr<SevenZipExtractFileOutStream> stream;
        const HRESULT hr = SevenZipExtractFileOutStream::Create(entry->targetPath, _conflictPolicy, stream);
        if (FAILED(hr))
        {
            if (_conflictPolicy == ArchiveExistingTargetPolicy::Skip && IsArchiveTargetExistsFailure(hr))
            {
                ++_result.skippedConflictCount;
                _activeEntry = nullptr;
                return S_OK;
            }
            _resultHr = hr;
            return hr;
        }

        _activeFileStream = stream;
        *outStream        = static_cast<ISequentialOutStream*>(stream.get());
        stream->AddRef();
        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE PrepareOperation(Int32) noexcept override
    {
        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE SetOperationResult(Int32 operationResult) noexcept override
    {
        if (operationResult != NArchive::NExtract::NOperationResult::kOK)
        {
            if (SUCCEEDED(_resultHr))
            {
                _resultHr = SevenZipExtractOperationResultToHr(operationResult);
            }
            _activeFileStream.reset();
            _activeEntry = nullptr;
            return S_OK;
        }

        if (_activeEntry)
        {
            if (_activeFileStream)
            {
                const uint64_t entryBytes = (_activeEntry->sizeBytes != 0u) ? _activeEntry->sizeBytes : _activeFileStream->BytesWritten();
                if (const HRESULT sizeHr = AccumulateArchiveExtractSize(_result.bytesProcessed, entryBytes); FAILED(sizeHr))
                {
                    _resultHr = sizeHr;
                    _activeFileStream.reset();
                    _activeEntry = nullptr;
                    return sizeHr;
                }
                const HRESULT commitHr = _activeFileStream->Commit();
                if (FAILED(commitHr))
                {
                    _resultHr = commitHr;
                    _activeFileStream.reset();
                    _activeEntry = nullptr;
                    return commitHr;
                }
            }

            _result.entries.push_back(_activeEntry->entryName);
            ++_result.entryCount;
        }

        _activeFileStream.reset();
        _activeEntry = nullptr;
        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE CryptoGetTextPassword(BSTR* password) noexcept override
    {
        if (! password)
        {
            return E_POINTER;
        }

        *password = nullptr;
        return HRESULT_FROM_WIN32(ERROR_INVALID_PASSWORD);
    }

    HRESULT STDMETHODCALLTYPE CryptoGetTextPassword2(Int32* passwordIsDefined, BSTR* password) noexcept override
    {
        if (! passwordIsDefined || ! password)
        {
            return E_POINTER;
        }

        *passwordIsDefined = 0;
        *password          = nullptr;
        return S_OK;
    }

private:
    [[nodiscard]] const SevenZipExtractEntry* FindEntry(UInt32 archiveIndex) const noexcept
    {
        for (const SevenZipExtractEntry& entry : _entries)
        {
            if (entry.archiveIndex == archiveIndex)
            {
                return &entry;
            }
        }
        return nullptr;
    }

    std::atomic_ulong _refCount{1u};
    const std::vector<SevenZipExtractEntry>& _entries;
    ArchiveExistingTargetPolicy _conflictPolicy = ArchiveExistingTargetPolicy::Skip;
    ArchiveOperationResult& _result;
    HRESULT _resultHr                        = S_OK;
    const SevenZipExtractEntry* _activeEntry = nullptr;
    wil::com_ptr<SevenZipExtractFileOutStream> _activeFileStream;
};

[[nodiscard]] std::wstring ArchiveStringPropertyForExtract(IInArchive* archive, UInt32 index, PROPID propId) noexcept
{
    if (! archive)
    {
        return {};
    }

    PROPVARIANT value{};
    PropVariantInit(&value);
    if (FAILED(archive->GetProperty(index, propId, &value)))
    {
        return {};
    }

    auto clearValue = wil::scope_exit([&] { PropVariantClear(&value); });
    return PropVariantToWideStringForArchive(value);
}

[[nodiscard]] bool ArchiveBoolPropertyForExtract(IInArchive* archive, UInt32 index, PROPID propId, bool& outValue) noexcept
{
    outValue = false;
    if (! archive)
    {
        return false;
    }

    PROPVARIANT value{};
    PropVariantInit(&value);
    if (FAILED(archive->GetProperty(index, propId, &value)))
    {
        return false;
    }

    auto clearValue = wil::scope_exit([&] { PropVariantClear(&value); });
    if (value.vt == VT_BOOL)
    {
        outValue = value.boolVal != VARIANT_FALSE;
        return true;
    }
    if (value.vt == VT_UI4)
    {
        outValue = value.ulVal != 0u;
        return true;
    }
    if (value.vt == VT_I4)
    {
        outValue = value.lVal != 0;
        return true;
    }
    return false;
}

[[nodiscard]] bool ArchiveLinkPropertyPresentForExtract(IInArchive* archive, UInt32 index, PROPID propId) noexcept
{
    if (! archive)
    {
        return false;
    }

    PROPVARIANT value{};
    PropVariantInit(&value);
    if (FAILED(archive->GetProperty(index, propId, &value)))
    {
        return false;
    }

    auto clearValue = wil::scope_exit([&] { PropVariantClear(&value); });
    switch (value.vt)
    {
        case VT_EMPTY:
        case VT_NULL: return false;
        case VT_BOOL: return value.boolVal != VARIANT_FALSE;
        case VT_UI1: return value.bVal != 0u;
        case VT_UI2: return value.uiVal != 0u;
        case VT_UI4: return value.ulVal != 0u;
        case VT_UI8: return value.uhVal.QuadPart != 0u;
        case VT_I1: return value.cVal != 0;
        case VT_I2: return value.iVal != 0;
        case VT_I4: return value.lVal != 0;
        case VT_I8: return value.hVal.QuadPart != 0;
        case VT_BSTR: return value.bstrVal && SysStringLen(value.bstrVal) > 0u;
        case VT_LPWSTR: return value.pwszVal && value.pwszVal[0] != L'\0';
        case VT_LPSTR: return value.pszVal && value.pszVal[0] != '\0';
        default: return true;
    }
}

[[nodiscard]] bool ArchiveUInt64PropertyForExtract(IInArchive* archive, UInt32 index, PROPID propId, uint64_t& outValue) noexcept
{
    outValue = 0u;
    if (! archive)
    {
        return false;
    }

    PROPVARIANT value{};
    PropVariantInit(&value);
    if (FAILED(archive->GetProperty(index, propId, &value)))
    {
        return false;
    }

    auto clearValue = wil::scope_exit([&] { PropVariantClear(&value); });
    return PropVariantToUInt64ForArchive(value, outValue);
}

[[nodiscard]] HRESULT BuildSevenZipExtractEntries(IInArchive* archive,
                                                  const std::filesystem::path& destinationPath,
                                                  std::wstring_view maskText,
                                                  std::vector<SevenZipExtractEntry>& outEntries)
{
    outEntries.clear();
    if (! archive)
    {
        return E_POINTER;
    }

    UInt32 itemCount = 0u;
    HRESULT hr       = archive->GetNumberOfItems(&itemCount);
    if (FAILED(hr))
    {
        return hr;
    }

    constexpr UInt32 kMaxArchiveExtractItems = 1'000'000u;
    if (itemCount > kMaxArchiveExtractItems)
    {
        return HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
    }

    const bool maskMatchesAll           = ArchiveUnpackMaskMatchesAll(maskText);
    const MaskSyntax::WildcardMask mask = maskMatchesAll ? MaskSyntax::WildcardMask{} : MaskSyntax::ParseWildcardMask(maskText);
    outEntries.reserve(itemCount);

    uint64_t totalUncompressedSize = 0u;
    for (UInt32 index = 0u; index < itemCount; ++index)
    {
        std::wstring entryName = ArchiveStringPropertyForExtract(archive, index, kpidPath);
        if (entryName.empty())
        {
            entryName = ArchiveStringPropertyForExtract(archive, index, kpidName);
        }
        if (entryName.empty())
        {
            continue;
        }

        if (ArchiveLinkPropertyPresentForExtract(archive, index, kpidSymLink) || ArchiveLinkPropertyPresentForExtract(archive, index, kpidHardLink))
        {
            return HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
        }

        bool isDirectory = false;
        if (! ArchiveBoolPropertyForExtract(archive, index, kpidIsDir, isDirectory))
        {
            isDirectory = ! entryName.empty() && (entryName.back() == L'/' || entryName.back() == L'\\');
        }
        if (isDirectory && entryName.back() != L'/')
        {
            entryName.push_back(L'/');
        }

        bool safeDirectory = false;
        std::filesystem::path relativePath;
        if (! IsArchiveEntryNameSafe(entryName, safeDirectory, relativePath))
        {
            return HRESULT_FROM_WIN32(ERROR_INVALID_NAME);
        }
        isDirectory = isDirectory || safeDirectory;

        if (! ShouldExtractArchiveEntryForMask(entryName, relativePath, mask, maskMatchesAll))
        {
            continue;
        }

        SevenZipExtractEntry entry{};
        entry.archiveIndex = index;
        entry.entryName    = std::move(entryName);
        entry.relativePath = std::move(relativePath);
        entry.targetPath   = destinationPath / entry.relativePath;
        if (! IsArchiveTargetPathInsideDestination(destinationPath, entry.targetPath))
        {
            return HRESULT_FROM_WIN32(ERROR_INVALID_NAME);
        }
        entry.directory = isDirectory;
        if (! entry.directory)
        {
            static_cast<void>(ArchiveUInt64PropertyForExtract(archive, index, kpidSize, entry.sizeBytes));
            if (const HRESULT sizeHr = AccumulateArchiveExtractSize(totalUncompressedSize, entry.sizeBytes); FAILED(sizeHr))
            {
                return sizeHr;
            }
        }
        outEntries.push_back(std::move(entry));
    }

    std::sort(outEntries.begin(),
              outEntries.end(),
              [](const SevenZipExtractEntry& left, const SevenZipExtractEntry& right) noexcept
    {
        if (left.directory != right.directory)
        {
            return left.directory;
        }
        return left.entryName < right.entryName;
    });
    return S_OK;
}

[[nodiscard]] ArchiveOperationResult ExtractSevenZipArchive(const std::filesystem::path& archivePath,
                                                            const std::filesystem::path& destinationPath,
                                                            ArchiveExistingTargetPolicy conflictPolicy,
                                                            std::wstring_view maskText)
{
    ArchiveOperationResult result{};
    result.archivePath     = archivePath;
    result.destinationPath = destinationPath;
    if (archivePath.empty() || destinationPath.empty())
    {
        result.hr = HRESULT_FROM_WIN32(ERROR_INVALID_NAME);
        return result;
    }

    std::error_code ec;
    std::filesystem::create_directories(destinationPath, ec);
    if (ec)
    {
        result.hr = HResultFromErrorCode(ec);
        return result;
    }

    SevenZipArchiveExports& api = GetSevenZipArchiveExports();
    result.hr                   = LoadSevenZipArchiveExports(api);
    if (FAILED(result.hr) || ! api.createObject)
    {
        return result;
    }

    const std::wstring extensionNoDotLower = ArchiveExtensionNoDotLower(archivePath);
    const std::optional<GUID> classId      = TryGetSevenZipArchiveClassIdForExtension(api, extensionNoDotLower);
    if (! classId.has_value())
    {
        result.hr = HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
        return result;
    }

    wil::com_ptr<IInArchive> archive;
    result.hr = api.createObject(&classId.value(), &IID_IInArchive, archive.put_void());
    if (FAILED(result.hr) || ! archive)
    {
        if (SUCCEEDED(result.hr))
        {
            result.hr = E_NOINTERFACE;
        }
        return result;
    }

    wil::com_ptr<IInStream> inputStream;
    result.hr = SevenZipExtractFileInStream::Create(archivePath, inputStream);
    if (FAILED(result.hr))
    {
        return result;
    }

    const HRESULT openHr = archive->Open(inputStream.get(), nullptr, nullptr);
    if (openHr != S_OK)
    {
        result.hr = FAILED(openHr) ? openHr : HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
        return result;
    }
    auto closeArchive = wil::scope_exit([&] { static_cast<void>(archive->Close()); });

    std::vector<SevenZipExtractEntry> entries;
    result.hr = BuildSevenZipExtractEntries(archive.get(), destinationPath, maskText, entries);
    if (FAILED(result.hr) || entries.empty())
    {
        return result;
    }

    std::vector<SevenZipExtractEntry> plannedEntries;
    plannedEntries.reserve(entries.size());
    for (SevenZipExtractEntry& entry : entries)
    {
        ArchiveTargetDecision decision = ArchiveTargetDecision::Extract;
        result.hr                      = ClassifyArchiveTarget(entry.targetPath, entry.directory, conflictPolicy, decision);
        if (FAILED(result.hr))
        {
            return result;
        }
        if (decision == ArchiveTargetDecision::Skip)
        {
            ++result.skippedConflictCount;
            continue;
        }
        plannedEntries.push_back(std::move(entry));
    }
    entries = std::move(plannedEntries);
    if (entries.empty())
    {
        return result;
    }

    std::vector<UInt32> indices;
    indices.reserve(entries.size());
    for (const SevenZipExtractEntry& entry : entries)
    {
        indices.push_back(entry.archiveIndex);
    }

    auto* callbackImpl = new (std::nothrow) SevenZipExtractCallback(entries, conflictPolicy, result);
    if (! callbackImpl)
    {
        result.hr = E_OUTOFMEMORY;
        return result;
    }

    wil::com_ptr<IArchiveExtractCallback> callback;
    callback.attach(callbackImpl);
    result.hr = archive->Extract(indices.data(), static_cast<UInt32>(indices.size()), 0, callback.get());
    if (SUCCEEDED(result.hr) && FAILED(callbackImpl->ResultHr()))
    {
        result.hr = callbackImpl->ResultHr();
    }
    return result;
}

[[nodiscard]] ArchiveOperationResult ExtractStoredZipArchive(const std::filesystem::path& archivePath,
                                                             const std::filesystem::path& destinationPath,
                                                             ArchiveExistingTargetPolicy conflictPolicy,
                                                             std::wstring_view maskText)
{
    ArchiveOperationResult result{};
    result.archivePath     = archivePath;
    result.destinationPath = destinationPath;
    if (archivePath.empty() || destinationPath.empty())
    {
        result.hr = HRESULT_FROM_WIN32(ERROR_INVALID_NAME);
        return result;
    }
    if (! PathMatchSpecW(archivePath.c_str(), L"*.zip"))
    {
        result.hr = HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
        return result;
    }

    std::error_code ec;
    std::filesystem::create_directories(destinationPath, ec);
    if (ec)
    {
        result.hr = HResultFromErrorCode(ec);
        return result;
    }

    const uint64_t fileSize = std::filesystem::file_size(archivePath, ec);
    if (ec)
    {
        result.hr = HResultFromErrorCode(ec);
        return result;
    }

#pragma warning(push)
#pragma warning(disable : 4625 4626) // WIL unique_hfile copy operations are intentionally deleted.
    wil::unique_hfile archiveFile(
        CreateFileW(archivePath.c_str(), GENERIC_READ, FILE_SHARE_READ, nullptr, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL | FILE_FLAG_SEQUENTIAL_SCAN, nullptr));
#pragma warning(pop)
    if (! archiveFile)
    {
        result.hr = HResultFromLastError();
        return result;
    }

    ZipEndOfCentralDirectory eocd{};
    result.hr = FindZipEndOfCentralDirectory(archiveFile.get(), fileSize, eocd);
    if (FAILED(result.hr))
    {
        return result;
    }

    std::vector<ZipParsedEntry> entries;
    result.hr = ParseZipCentralDirectory(archiveFile.get(), eocd, entries);
    if (FAILED(result.hr))
    {
        return result;
    }

    std::sort(entries.begin(),
              entries.end(),
              [](const ZipParsedEntry& left, const ZipParsedEntry& right) noexcept
    {
        if (left.directory != right.directory)
        {
            return left.directory;
        }
        return left.entryName < right.entryName;
    });

    const bool maskMatchesAll           = ArchiveUnpackMaskMatchesAll(maskText);
    const MaskSyntax::WildcardMask mask = maskMatchesAll ? MaskSyntax::WildcardMask{} : MaskSyntax::ParseWildcardMask(maskText);
    std::vector<const ZipParsedEntry*> plannedEntries;
    plannedEntries.reserve(entries.size());
    for (const ZipParsedEntry& entry : entries)
    {
        if (! ShouldExtractZipEntryForMask(entry, mask, maskMatchesAll))
        {
            continue;
        }

        const std::filesystem::path targetPath = destinationPath / entry.relativePath;
        if (! IsArchiveTargetPathInsideDestination(destinationPath, targetPath))
        {
            result.hr = HRESULT_FROM_WIN32(ERROR_INVALID_NAME);
            return result;
        }
        ArchiveTargetDecision decision = ArchiveTargetDecision::Extract;
        result.hr                      = ClassifyArchiveTarget(targetPath, entry.directory, conflictPolicy, decision);
        if (FAILED(result.hr))
        {
            return result;
        }
        if (decision == ArchiveTargetDecision::Skip)
        {
            ++result.skippedConflictCount;
            continue;
        }
        plannedEntries.push_back(&entry);
    }

    for (const ZipParsedEntry* entryPointer : plannedEntries)
    {
        const ZipParsedEntry& entry            = *entryPointer;
        const std::filesystem::path targetPath = destinationPath / entry.relativePath;
        if (entry.directory)
        {
            std::filesystem::create_directories(targetPath, ec);
            if (ec)
            {
                result.hr = HResultFromErrorCode(ec);
                return result;
            }
        }
        else
        {
            uint64_t dataOffset = 0u;
            result.hr           = ResolveZipEntryDataOffset(archiveFile.get(), entry, fileSize, dataOffset);
            if (FAILED(result.hr))
            {
                return result;
            }

            uint32_t crc32 = 0u;
            result.hr      = ExtractStoredZipFileEntry(archiveFile.get(), entry, dataOffset, targetPath, conflictPolicy, crc32);
            if (result.hr == S_FALSE)
            {
                ++result.skippedConflictCount;
                result.hr = S_OK;
                continue;
            }
            if (FAILED(result.hr))
            {
                return result;
            }
            result.bytesProcessed += entry.uncompressedSize;
        }

        result.entries.push_back(entry.entryName);
        ++result.entryCount;
    }

    return result;
}

void ShowArchiveOverlay(
    FolderWindow& window, FolderWindow::Pane pane, UINT titleStringId, FolderView::OverlaySeverity severity, std::wstring message, HRESULT hr) noexcept
{
    Debug::Perf::Scope perf(L"archive.feedback_us");
    perf.SetHr(hr);

    std::wstring title = LoadStringResource(nullptr, titleStringId);
    if (title.empty())
    {
        title = LoadStringResource(nullptr, severity == FolderView::OverlaySeverity::Information ? IDS_OVERLAY_TITLE_INFORMATION : IDS_CAPTION_WARNING);
    }

    window.ShowPaneAlertOverlay(pane,
                                FolderView::ErrorOverlayKind::Operation,
                                severity,
                                std::move(title),
                                std::move(message),
                                hr,
                                severity != FolderView::OverlaySeverity::Information,
                                false);
}

[[nodiscard]] std::wstring ArchiveFailureMessage(UINT formatStringId, const std::filesystem::path& archivePath, HRESULT hr)
{
    std::wstring message = FormatStringResource(nullptr, formatStringId, archivePath.wstring(), static_cast<unsigned long>(static_cast<uint32_t>(hr)));
    if (message.empty())
    {
        message = std::format(L"Archive operation failed for {} (0x{:08X}).", archivePath.wstring(), static_cast<unsigned long>(static_cast<uint32_t>(hr)));
    }
    return message;
}

void RefreshFolderViewIfPathMatches(FolderView& folderView, const std::filesystem::path& path)
{
    const std::optional<std::filesystem::path> currentFolder = folderView.GetFolderPath();
    if (! currentFolder.has_value() || path.empty())
    {
        return;
    }

    std::error_code ec;
    if (std::filesystem::equivalent(currentFolder.value(), path, ec) && ! ec)
    {
        folderView.ForceRefresh();
    }
}

[[nodiscard]] bool ConfirmPermanentDeletePaths(HWND ownerWindow, const std::vector<std::filesystem::path>& sourcePaths) noexcept
{
    if (sourcePaths.empty())
    {
        return true;
    }

    auto suffixFor = [](unsigned long long count) noexcept -> std::wstring_view { return count == 1ull ? std::wstring_view(L"") : std::wstring_view(L"s"); };

    auto ensureTrailingSeparator = [](std::wstring text) noexcept -> std::wstring
    {
        if (text.empty())
        {
            return text;
        }

        const wchar_t last = text.back();
        if (last == L'\\' || last == L'/')
        {
            return text;
        }

        text.push_back(L'\\');
        return text;
    };

    auto normalizeSlashes = [](std::wstring& text) noexcept
    {
        for (auto& ch : text)
        {
            if (ch == L'/')
            {
                ch = L'\\';
            }
        }
    };

    const unsigned long long itemCount = static_cast<unsigned long long>(sourcePaths.size());
    const std::wstring_view itemSuffix = suffixFor(itemCount);
    const std::wstring what            = FormatStringResource(nullptr, IDS_FMT_FILEOPS_COUNT_ITEM, itemCount, itemSuffix);

    std::wstring fromText;
    if (sourcePaths.size() == 1u)
    {
        fromText = sourcePaths.front().wstring();
    }
    else
    {
        std::filesystem::path commonParent = sourcePaths.front().parent_path();
        bool multipleParents               = false;
        for (size_t index = 1; index < sourcePaths.size(); ++index)
        {
            const std::filesystem::path parent = sourcePaths[index].parent_path();
            if (CompareStringOrdinal(commonParent.c_str(), -1, parent.c_str(), -1, TRUE) != CSTR_EQUAL)
            {
                multipleParents = true;
                break;
            }
        }

        fromText = multipleParents ? LoadStringResource(nullptr, IDS_FILEOPS_LOCATION_MULTIPLE) : ensureTrailingSeparator(commonParent.wstring());
    }

    normalizeSlashes(fromText);

    const std::wstring message = FormatStringResource(nullptr, IDS_FMT_FILEOPS_CONFIRM_PERMANENT_DELETE, what, fromText);
    const std::wstring caption = LoadStringResource(nullptr, IDS_CAPTION_CONFIRM);

    HostPromptRequest prompt{};
    prompt.version       = 1;
    prompt.sizeBytes     = sizeof(prompt);
    prompt.scope         = (ownerWindow && IsWindow(ownerWindow)) ? HOST_ALERT_SCOPE_WINDOW : HOST_ALERT_SCOPE_APPLICATION;
    prompt.severity      = HOST_ALERT_WARNING;
    prompt.buttons       = HOST_PROMPT_BUTTONS_OK_CANCEL;
    prompt.targetWindow  = prompt.scope == HOST_ALERT_SCOPE_WINDOW ? ownerWindow : nullptr;
    prompt.title         = caption.c_str();
    prompt.message       = message.c_str();
    prompt.defaultResult = HOST_PROMPT_RESULT_CANCEL;

    HostPromptResult promptResult = HOST_PROMPT_RESULT_NONE;
    const HRESULT hrPrompt        = HostShowPrompt(prompt, nullptr, &promptResult);
    return SUCCEEDED(hrPrompt) && promptResult == HOST_PROMPT_RESULT_OK;
}

[[nodiscard]] HRESULT DeleteUnpackedArchives(const std::vector<std::filesystem::path>& archivePaths) noexcept
{
    for (const std::filesystem::path& path : archivePaths)
    {
        std::error_code ec;
        const bool exists = std::filesystem::exists(path, ec);
        if (ec)
        {
            return HResultFromErrorCode(ec);
        }
        if (! exists)
        {
            continue;
        }

        static_cast<void>(std::filesystem::remove(path, ec));
        if (ec)
        {
            return HResultFromErrorCode(ec);
        }
    }

    return S_OK;
}

[[nodiscard]] std::wstring OpenedFilesDisplayNameForPath(const std::filesystem::path& path)
{
    std::wstring name = path.filename().wstring();
    if (name.empty())
    {
        name = path.native();
    }
    return name;
}

[[nodiscard]] bool SharedDirectoryPathExists(std::wstring_view localPath) noexcept
{
    if (localPath.empty())
    {
        return false;
    }

    std::error_code ec;
    return std::filesystem::is_directory(std::filesystem::path(localPath), ec);
}

} // namespace

LRESULT FolderWindow::OnMakeFileListTaskUpdate(LPARAM lp) noexcept
{
    auto payload = TakeMessagePayload<MakeFileListTaskPayload>(lp);
    if (! payload)
    {
        return 0;
    }

    return static_cast<LRESULT>(CreateOrUpdateInformationalTask(payload->update));
}

void FolderWindow::RequestMakeFileListCancellation(const Pane pane) noexcept
{
    PaneState& state = pane == Pane::Left ? _leftPane : _rightPane;
    if (! state.makeFileListThread.joinable())
    {
        return;
    }

    state.makeFileListThread.request_stop();
    if (CancelSynchronousIo(state.makeFileListThread.native_handle()) == FALSE)
    {
        const DWORD error = GetLastError();
        if (error != ERROR_NOT_FOUND)
        {
            Debug::Warning(L"Make File List: CancelSynchronousIo failed (gle=0x{:08X})", error);
        }
    }
}

LRESULT FolderWindow::OnMakeFileListCompleted(LPARAM lp) noexcept
{
    auto payload = TakeMessagePayload<MakeFileListCompletedPayload>(lp);
    if (! payload)
    {
        return 0;
    }

    PaneState& state = payload->pane == Pane::Left ? _leftPane : _rightPane;
    if (state.makeFileListThread.joinable())
    {
        state.makeFileListThread = {};
    }

    HRESULT hr                 = payload->hr;
    const HRESULT cancelledHr  = HRESULT_FROM_WIN32(ERROR_CANCELLED);
    const auto outputStartedAt = std::chrono::steady_clock::now();
    if (SUCCEEDED(hr) && payload->options.outputTarget == Common::Settings::MakeFileListOutputTarget::Clipboard)
    {
        hr = Common::Clipboard::TrySetUnicodeText(GetClipboardOwnerWindow(_hWnd.get()), payload->clipboardText) ? S_OK
                                                                                                                : HRESULT_FROM_WIN32(ERROR_CLIPBOARD_NOT_OPEN);
    }

    Debug::Perf::Emit(L"makeFileList.output_us",
                      payload->options.outputTarget == Common::Settings::MakeFileListOutputTarget::File ? L"file" : L"clipboard",
                      payload->options.outputTarget == Common::Settings::MakeFileListOutputTarget::File ? payload->outputElapsedUs
                                                                                                        : Debug::Perf::ElapsedUs(outputStartedAt),
                      payload->outputBytes,
                      payload->entryCount,
                      hr);

    InformationalTaskUpdate finalTask{};
    finalTask.kind                        = InformationalTaskUpdate::Kind::MakeFileList;
    finalTask.taskId                      = payload->taskId;
    finalTask.title                       = payload->title;
    finalTask.makeFileListCurrentPath     = payload->currentFolder;
    finalTask.makeFileListScannedEntries  = payload->entryCount;
    finalTask.makeFileListTotalEntries    = payload->entryCount;
    finalTask.makeFileListRenderedEntries = payload->entryCount;
    finalTask.finished                    = true;
    finalTask.resultHr                    = hr;

    if (hr == cancelledHr || hr == E_ABORT)
    {
        finalTask.doneSummary = LoadStringResource(nullptr, IDS_OVERLAY_MSG_ENUMERATION_CANCELED);
        static_cast<void>(CreateOrUpdateInformationalTask(finalTask));
        ShowMakeFileListOverlay(*this, payload->pane, FolderView::OverlaySeverity::Information, finalTask.doneSummary, hr);
    }
    else if (FAILED(hr))
    {
        finalTask.doneSummary =
            FormatStringResource(nullptr,
                                 IDS_FMT_MAKE_FILE_LIST_FAILED,
                                 payload->options.outputTarget == Common::Settings::MakeFileListOutputTarget::File ? payload->options.outputFile.wstring()
                                                                                                                   : payload->currentFolder.wstring(),
                                 static_cast<unsigned long>(static_cast<uint32_t>(hr)));
        static_cast<void>(CreateOrUpdateInformationalTask(finalTask));

        const std::wstring message = payload->options.outputTarget == Common::Settings::MakeFileListOutputTarget::File
                                         ? finalTask.doneSummary
                                         : LoadStringResource(nullptr, IDS_MSG_SELECTION_SAVE_CLIPBOARD_FAILED);
        ShowMakeFileListOverlay(*this, payload->pane, FolderView::OverlaySeverity::Error, message, hr);
    }
    else
    {
        if (_settings)
        {
            _settings->makeFileList = payload->options;
        }

        finalTask.doneSummary = FormatStringResource(nullptr,
                                                     IDS_FMT_MAKE_FILE_LIST_COMPLETED,
                                                     MakeFileListFormatDetail(payload->options.format),
                                                     static_cast<unsigned long long>(payload->entryCount),
                                                     payload->outputTarget);
        static_cast<void>(CreateOrUpdateInformationalTask(finalTask));
        ShowMakeFileListOverlay(*this, payload->pane, FolderView::OverlaySeverity::Information, finalTask.doneSummary, S_OK);
    }

    Debug::Perf::Emit(L"makeFileList.total_us",
                      MakeFileListFormatDetail(payload->options.format),
                      payload->totalElapsedUs,
                      payload->entryCount,
                      payload->collectFailures,
                      FAILED(hr) ? hr : (payload->collectFailures == 0u ? S_OK : S_FALSE));
    return 0;
}

LRESULT FolderWindow::OnChangeAttributesTaskUpdate(LPARAM lp) noexcept
{
    auto payload = TakeMessagePayload<ChangeAttributesTaskPayload>(lp);
    if (! payload)
    {
        return 0;
    }

    return static_cast<LRESULT>(CreateOrUpdateInformationalTask(payload->update));
}

LRESULT FolderWindow::OnChangeAttributesCompleted(LPARAM lp) noexcept
{
    auto payload = TakeMessagePayload<ChangeAttributesCompletedPayload>(lp);
    if (! payload)
    {
        return 0;
    }

    PaneState& state = payload->pane == Pane::Left ? _leftPane : _rightPane;
    if (state.changeAttributesThread.joinable())
    {
        state.changeAttributesThread = {};
    }

#ifdef ENABLE_TESTS
    _debugLastChangeAttributesReport = payload->report;
#endif

    if (payload->refreshNeeded)
    {
        state.folderView.ForceRefresh();
    }

    ShowChangeAttributesReportOverlay(*this, payload->pane, payload->report);
    return 0;
}

struct FolderWindow::OpenedFilesDialogState final : RedSalamander::DxUi::IDxGridModel, RedSalamander::DxUi::IDxGridDelegate
{
    using RedSalamander::DxUi::IDxGridDelegate::OnGridRowActivated;
    using RedSalamander::DxUi::IDxGridDelegate::OnGridSelectionChanged;

    OpenedFilesDialogState()                                         = default;
    OpenedFilesDialogState(const OpenedFilesDialogState&)            = delete;
    OpenedFilesDialogState& operator=(const OpenedFilesDialogState&) = delete;
    OpenedFilesDialogState(OpenedFilesDialogState&&)                 = delete;
    OpenedFilesDialogState& operator=(OpenedFilesDialogState&&)      = delete;

    static HRESULT EnsureWindowClass() noexcept
    {
        static ATOM atom = 0;
        if (atom != 0)
        {
            return S_OK;
        }

        WNDCLASSEXW wc{};
        wc.cbSize        = sizeof(wc);
        wc.lpfnWndProc   = OpenedFilesDialogState::WndProc;
        wc.hInstance     = GetModuleHandleW(nullptr);
        wc.hCursor       = LoadCursorW(nullptr, IDC_ARROW);
        wc.lpszClassName = kOpenedFilesWindowClassName;
        wc.style         = CS_DBLCLKS;

        atom = RegisterClassExW(&wc);
        return atom != 0 ? S_OK : HRESULT_FROM_WIN32(GetLastError());
    }

    static LRESULT CALLBACK WndProc(HWND hwnd, UINT message, WPARAM wParam, LPARAM lParam) noexcept
    {
        if (message == WM_NCCREATE)
        {
            auto* cs    = reinterpret_cast<CREATESTRUCTW*>(lParam);
            auto* state = static_cast<OpenedFilesDialogState*>(cs ? cs->lpCreateParams : nullptr);
            if (! state)
            {
                return FALSE;
            }

            SetWindowLongPtrW(hwnd, GWLP_USERDATA, reinterpret_cast<LONG_PTR>(state));
            if (! state->hwnd)
            {
                state->hwnd.reset(hwnd);
            }
            return TRUE;
        }

        auto* state = reinterpret_cast<OpenedFilesDialogState*>(GetWindowLongPtrW(hwnd, GWLP_USERDATA));
        if (! state)
        {
            return DefWindowProcW(hwnd, message, wParam, lParam);
        }

        bool handled     = false;
        LRESULT dxResult = 0;
        if (message != WM_CREATE)
        {
            dxResult = state->dxHost.HandleMessage(hwnd, message, wParam, lParam, handled);
        }
        if (handled)
        {
            if (message == WM_SIZE || message == WM_DPICHANGED)
            {
                state->Layout();
            }
            if (message == WM_NCDESTROY)
            {
                state->OnNcDestroy(hwnd);
            }
            return dxResult;
        }

        switch (message)
        {
            case WM_CREATE: return state->OnCreate(hwnd) ? 0 : -1;
            case WM_SIZE: state->Layout(); return 0;
            case WM_DPICHANGED:
            {
                const auto* suggested = reinterpret_cast<const RECT*>(lParam);
                if (suggested)
                {
                    SetWindowPos(hwnd,
                                 nullptr,
                                 suggested->left,
                                 suggested->top,
                                 suggested->right - suggested->left,
                                 suggested->bottom - suggested->top,
                                 SWP_NOZORDER | SWP_NOACTIVATE);
                }
                state->Layout();
                return 0;
            }
            case WM_GETMINMAXINFO:
            {
                auto* info = reinterpret_cast<MINMAXINFO*>(lParam);
                if (info)
                {
                    const UINT dpi         = GetDpiForWindow(hwnd);
                    info->ptMinTrackSize.x = ScalePanePromptForDpi(dpi, 440);
                    info->ptMinTrackSize.y = ScalePanePromptForDpi(dpi, 260);
                }
                return 0;
            }
            case WM_COMMAND:
                switch (LOWORD(wParam))
                {
                    case IDOK:
                    case IDC_OPENED_FILES_FOCUS:
                        if (state->owner)
                        {
                            static_cast<void>(state->owner->FocusOpenedFilesDialogSelection());
                        }
                        return 0;
                    case IDCANCEL:
                        if (state->owner)
                        {
                            state->owner->RequestCloseOpenedFilesDialog();
                        }
                        return 0;
                }
                break;
            case WM_ERASEBKGND: return 1;
            case WM_NCACTIVATE: ApplyTitleBarTheme(hwnd, state->theme, wParam != FALSE); return DefWindowProcW(hwnd, message, wParam, lParam);
            case WM_CLOSE:
                if (state->owner)
                {
                    state->owner->RequestCloseOpenedFilesDialog();
                    return 0;
                }
                break;
            case WM_NCDESTROY: state->OnNcDestroy(hwnd); break;
        }

        return DefWindowProcW(hwnd, message, wParam, lParam);
    }

    [[nodiscard]] size_t GetRowCount() const noexcept override
    {
        return rows.size();
    }

    [[nodiscard]] size_t GetColumnCount() const noexcept override
    {
        return columns.size();
    }

    [[nodiscard]] RedSalamander::DxUi::GridColumnDesc GetColumn(size_t columnIndex) const override
    {
        return columns.at(columnIndex);
    }

    void GetCellData(size_t rowIndex, size_t columnIndex, RedSalamander::DxUi::GridCellData& outCell) const override
    {
        outCell = {};
        if (rowIndex >= rows.size() || columnIndex >= columns.size())
        {
            return;
        }

        const OpenedFileRow& row = rows[rowIndex];
        switch (columnIndex)
        {
            case 0u: outCell.text = row.file; break;
            case 1u: outCell.text = row.source; break;
            case 2u: outCell.text = row.openedBy; break;
        }
        outCell.tooltipText = row.path.wstring();
    }

    [[nodiscard]] uint64_t GetStableRowId(size_t rowIndex) const noexcept override
    {
        return rowIndex < rows.size() ? static_cast<uint64_t>(rowIndex + 1u) : 0u;
    }

    [[nodiscard]] std::optional<size_t> FindRowByStableId(uint64_t rowId) const noexcept override
    {
        if (rowId == 0u)
        {
            return std::nullopt;
        }
        const size_t rowIndex = static_cast<size_t>(rowId - 1u);
        return rowIndex < rows.size() ? std::optional<size_t>(rowIndex) : std::nullopt;
    }

    void OnGridSelectionChanged(RedSalamander::DxUi::Grid& sender) override
    {
        selectedIndex = kOpenedFilesNoSelection;
        for (size_t rowIndex = 0u; rowIndex < rows.size(); ++rowIndex)
        {
            if (sender.GetSelectionModel().IsSelected(GetStableRowId(rowIndex)))
            {
                selectedIndex = rowIndex;
                break;
            }
        }
        UpdateEmptyState();
    }

    void OnGridRowActivated(RedSalamander::DxUi::Grid& /*sender*/, size_t rowIndex) override
    {
        if (rowIndex < rows.size() && hwnd && IsWindow(hwnd.get()) != FALSE)
        {
            selectedIndex = rowIndex;
            PostMessageW(hwnd.get(), WM_COMMAND, MAKEWPARAM(IDOK, BN_CLICKED), 0);
        }
    }

    void SetRows(std::vector<OpenedFileRow> newRows)
    {
        rows          = std::move(newRows);
        selectedIndex = rows.empty() ? kOpenedFilesNoSelection : 0u;
        if (grid)
        {
            grid->NotifyDataChanged();
            if (rows.empty())
            {
                grid->GetSelectionModel().Clear();
            }
            else
            {
                grid->GetSelectionModel().SetSingle(GetStableRowId(selectedIndex));
                grid->EnsureRowVisible(selectedIndex);
            }
        }
        UpdateEmptyState();
    }

    [[nodiscard]] bool SelectRow(size_t rowIndex) noexcept
    {
        if (! grid || rowIndex >= rows.size())
        {
            return false;
        }
        selectedIndex = rowIndex;
        grid->GetSelectionModel().SetSingle(GetStableRowId(rowIndex));
        grid->EnsureRowVisible(rowIndex);
        UpdateEmptyState();
        if (hwnd && IsWindow(hwnd.get()) != FALSE)
        {
            InvalidateRect(hwnd.get(), nullptr, FALSE);
        }
        return true;
    }

    [[nodiscard]] bool OnCreate(HWND createdHwnd) noexcept
    {
        if (! dxHost.Attach(createdHwnd))
        {
            return false;
        }
        BuildUi();
        ApplyTheme();
        Layout();
        if (grid && ! rows.empty())
        {
            dxHost.SetFocusControl(grid);
        }
        else if (closeButton)
        {
            dxHost.SetFocusControl(closeButton);
        }
        dxHost.SetDefaultButton(focusButton);
        dxHost.SetCancelButton(closeButton);
        return true;
    }

    void BuildUi()
    {
        if (root)
        {
            return;
        }

        using namespace RedSalamander::DxUi;

        columns = {
            {L"file", LoadStringResource(nullptr, IDS_OPENED_FILES_COLUMN_FILE), 260.0f, 140.0f, GridColumnKind::Text, false, false},
            {L"source", LoadStringResource(nullptr, IDS_OPENED_FILES_COLUMN_SOURCE), 130.0f, 96.0f, GridColumnKind::Text, false, false},
            {L"openedBy", LoadStringResource(nullptr, IDS_OPENED_FILES_COLUMN_OPENED_BY), 170.0f, 120.0f, GridColumnKind::Text, false, false},
        };

        auto rootOwned = std::make_unique<Panel>();
        root           = rootOwned.get();

        grid = root->AddChild<Grid>();
        grid->SetModel(this);
        grid->SetDelegate(this);
        grid->SetSelectionMode(GridSelectionMode::Single);
        grid->SetEmptyStateText(LoadStringOrFallback(IDS_OPENED_FILES_EMPTY, L"No files are currently opened by RedSalamander."));
        grid->SetLineClamp(1u);

        emptyLabel = root->AddChild<Label>(LoadStringOrFallback(IDS_OPENED_FILES_EMPTY, L"No files are currently opened by RedSalamander."));
        emptyLabel->SetAlignment(DWRITE_TEXT_ALIGNMENT_LEADING);
        emptyLabel->SetMultiline(true);

        focusButton = root->AddChild<Button>(LoadStringOrFallback(IDS_OPENED_FILES_FOCUS_ITEM, L"Focus Item"));
        focusButton->SetPrimary(true);
        focusButton->SetOnClick([this]
        {
            if (hwnd && IsWindow(hwnd.get()) != FALSE)
            {
                PostMessageW(hwnd.get(), WM_COMMAND, MAKEWPARAM(IDC_OPENED_FILES_FOCUS, BN_CLICKED), 0);
            }
        });

        closeButton = root->AddChild<Button>(LoadStringOrFallback(IDS_OPENED_FILES_CLOSE, L"Close"));
        closeButton->SetOnClick([this]
        {
            if (hwnd && IsWindow(hwnd.get()) != FALSE)
            {
                PostMessageW(hwnd.get(), WM_COMMAND, MAKEWPARAM(IDCANCEL, BN_CLICKED), 0);
            }
        });

        dxHost.SetRoot(std::move(rootOwned));
        SetRows(std::move(rows));
    }

    void ApplyTheme() noexcept
    {
        palette = MakeAppThemeDxPalette(theme, theme.windowBackground);
        dxHost.SetTheme(palette);
        if (hwnd && IsWindow(hwnd.get()) != FALSE)
        {
            ApplyWindowChromeTheme(hwnd.get(), theme, WindowBackdropTarget::Tool, GetActiveWindow() == hwnd.get());
            RedrawWindow(hwnd.get(), nullptr, nullptr, RDW_FRAME | RDW_INVALIDATE | RDW_UPDATENOW);
        }
    }

    void Layout() noexcept
    {
        if (! root)
        {
            return;
        }

        const D2D1_RECT_F client = dxHost.GetClientBoundsDip();
        root->SetBounds(client);

        constexpr float kMarginDip       = 20.0f;
        constexpr float kGapDip          = 12.0f;
        constexpr float kButtonWidthDip  = 128.0f;
        constexpr float kButtonHeightDip = 34.0f;

        const float left       = client.left + kMarginDip;
        const float right      = std::max(left, client.right - kMarginDip);
        const float bottom     = std::max(client.top, client.bottom - kMarginDip);
        const float buttonTop  = std::max(client.top + kMarginDip, bottom - kButtonHeightDip);
        const float contentBot = std::max(client.top + kMarginDip, buttonTop - kGapDip);

        if (grid)
        {
            grid->SetBounds(D2D1::RectF(left, client.top + kMarginDip, right, contentBot));
        }
        if (emptyLabel)
        {
            emptyLabel->SetBounds(D2D1::RectF(left, client.top + kMarginDip, right, std::min(contentBot, client.top + kMarginDip + 72.0f)));
        }

        float x = right;
        if (closeButton)
        {
            closeButton->SetBounds(D2D1::RectF(x - kButtonWidthDip, buttonTop, x, buttonTop + kButtonHeightDip));
            x -= kButtonWidthDip + kGapDip;
        }
        if (focusButton)
        {
            focusButton->SetBounds(D2D1::RectF(x - kButtonWidthDip, buttonTop, x, buttonTop + kButtonHeightDip));
        }
    }

    void UpdateEmptyState() noexcept
    {
        const bool empty = rows.empty();
        if (grid)
        {
            grid->SetVisible(! empty);
        }
        if (emptyLabel)
        {
            emptyLabel->SetVisible(empty);
        }
        if (focusButton)
        {
            focusButton->SetEnabled(! empty && selectedIndex != kOpenedFilesNoSelection && selectedIndex < rows.size());
        }
    }

    void OnNcDestroy(HWND destroyedHwnd) noexcept
    {
        destroyed = true;
        dxHost.Detach();
        if (hwnd.get() == destroyedHwnd)
        {
            static_cast<void>(hwnd.release());
        }
        SetWindowLongPtrW(destroyedHwnd, GWLP_USERDATA, 0);
    }

    FolderWindow* owner = nullptr;
    Pane pane           = Pane::Left;
    AppTheme theme;
    RedSalamander::DxUi::ThemePalette palette{};
    RedSalamander::DxUi::WindowHost dxHost;
    RedSalamander::DxUi::Panel* root         = nullptr;
    RedSalamander::DxUi::Grid* grid          = nullptr;
    RedSalamander::DxUi::Label* emptyLabel   = nullptr;
    RedSalamander::DxUi::Button* focusButton = nullptr;
    RedSalamander::DxUi::Button* closeButton = nullptr;
    std::vector<RedSalamander::DxUi::GridColumnDesc> columns;
    std::vector<OpenedFileRow> rows;
    size_t selectedIndex = kOpenedFilesNoSelection;
    bool destroyed       = false;
    wil::unique_hwnd hwnd;
};

void FolderWindow::OpenedFilesDialogStateDeleter::operator()(OpenedFilesDialogState* state) const noexcept
{
    delete state;
}

HRESULT FolderWindow::SharedDirectoriesDialogState::EnsureWindowClass() noexcept
{
    static ATOM atom = 0;
    if (atom != 0)
    {
        return S_OK;
    }

    WNDCLASSEXW wc{};
    wc.cbSize        = sizeof(wc);
    wc.lpfnWndProc   = SharedDirectoriesDialogState::WndProc;
    wc.hInstance     = GetModuleHandleW(nullptr);
    wc.hCursor       = LoadCursorW(nullptr, IDC_ARROW);
    wc.lpszClassName = kSharedDirectoriesWindowClassName;
    wc.style         = CS_DBLCLKS;

    atom = RegisterClassExW(&wc);
    return atom != 0 ? S_OK : HRESULT_FROM_WIN32(GetLastError());
}

LRESULT CALLBACK FolderWindow::SharedDirectoriesDialogState::WndProc(HWND hwnd, UINT message, WPARAM wParam, LPARAM lParam) noexcept
{
    if (message == WM_NCCREATE)
    {
        auto* cs    = reinterpret_cast<CREATESTRUCTW*>(lParam);
        auto* state = static_cast<SharedDirectoriesDialogState*>(cs ? cs->lpCreateParams : nullptr);
        if (! state)
        {
            return FALSE;
        }

        SetWindowLongPtrW(hwnd, GWLP_USERDATA, reinterpret_cast<LONG_PTR>(state));
        if (! state->hwnd)
        {
            state->hwnd.reset(hwnd);
        }
        return TRUE;
    }

    auto* state = reinterpret_cast<SharedDirectoriesDialogState*>(GetWindowLongPtrW(hwnd, GWLP_USERDATA));
    if (! state)
    {
        return DefWindowProcW(hwnd, message, wParam, lParam);
    }

    bool handled     = false;
    LRESULT dxResult = 0;
    if (message != WM_CREATE)
    {
        dxResult = state->dxHost.HandleMessage(hwnd, message, wParam, lParam, handled);
    }
    if (handled)
    {
        if (message == WM_SIZE || message == WM_DPICHANGED)
        {
            state->Layout();
        }
        if (message == WM_NCDESTROY)
        {
            state->OnNcDestroy(hwnd);
        }
        return dxResult;
    }

    switch (message)
    {
        case WM_CREATE: return state->OnCreate(hwnd) ? 0 : -1;
        case WM_SIZE: state->Layout(); return 0;
        case WM_DPICHANGED:
        {
            const auto* suggested = reinterpret_cast<const RECT*>(lParam);
            if (suggested)
            {
                SetWindowPos(hwnd,
                             nullptr,
                             suggested->left,
                             suggested->top,
                             suggested->right - suggested->left,
                             suggested->bottom - suggested->top,
                             SWP_NOZORDER | SWP_NOACTIVATE);
            }
            state->Layout();
            return 0;
        }
        case WM_GETMINMAXINFO:
        {
            auto* info = reinterpret_cast<MINMAXINFO*>(lParam);
            if (info)
            {
                const UINT dpi         = GetDpiForWindow(hwnd);
                info->ptMinTrackSize.x = ScalePanePromptForDpi(dpi, 520);
                info->ptMinTrackSize.y = ScalePanePromptForDpi(dpi, 300);
            }
            return 0;
        }
        case WM_COMMAND:
            switch (LOWORD(wParam))
            {
                case IDOK:
                case IDC_SHARED_DIRECTORIES_OPEN:
                    if (state->owner)
                    {
                        static_cast<void>(state->owner->OpenSharedDirectoriesDialogSelection());
                    }
                    return 0;
                case IDC_SHARED_DIRECTORIES_MANAGE:
                    if (state->owner)
                    {
                        state->owner->OpenSharedDirectoriesManagement();
                    }
                    return 0;
                case IDCANCEL:
                    if (state->owner)
                    {
                        state->owner->RequestCloseSharedDirectoriesDialog();
                    }
                    return 0;
            }
            break;
        case WM_ERASEBKGND: return 1;
        case WM_NCACTIVATE: ApplyTitleBarTheme(hwnd, state->theme, wParam != FALSE); return DefWindowProcW(hwnd, message, wParam, lParam);
        case WM_CLOSE:
            if (state->owner)
            {
                state->owner->RequestCloseSharedDirectoriesDialog();
                return 0;
            }
            break;
        case WM_NCDESTROY: state->OnNcDestroy(hwnd); break;
    }

    return DefWindowProcW(hwnd, message, wParam, lParam);
}

size_t FolderWindow::SharedDirectoriesDialogState::GetRowCount() const noexcept
{
    return rows.size();
}

size_t FolderWindow::SharedDirectoriesDialogState::GetColumnCount() const noexcept
{
    return columns.size();
}

RedSalamander::DxUi::GridColumnDesc FolderWindow::SharedDirectoriesDialogState::GetColumn(size_t columnIndex) const
{
    return columns.at(columnIndex);
}

void FolderWindow::SharedDirectoriesDialogState::GetCellData(size_t rowIndex, size_t columnIndex, RedSalamander::DxUi::GridCellData& outCell) const
{
    outCell = {};
    if (rowIndex >= rows.size() || columnIndex >= columns.size())
    {
        return;
    }

    const SharedDirectoryRow& row = rows[rowIndex];
    switch (columnIndex)
    {
        case 0u: outCell.text = row.name; break;
        case 1u: outCell.text = row.localPath; break;
        case 2u: outCell.text = row.type; break;
        case 3u: outCell.text = row.remark; break;
    }
    outCell.tooltipText = row.localPath.empty() ? row.remark : row.localPath;
}

uint64_t FolderWindow::SharedDirectoriesDialogState::GetStableRowId(size_t rowIndex) const noexcept
{
    return rowIndex < rows.size() ? static_cast<uint64_t>(rowIndex + 1u) : 0u;
}

std::optional<size_t> FolderWindow::SharedDirectoriesDialogState::FindRowByStableId(uint64_t rowId) const noexcept
{
    if (rowId == 0u)
    {
        return std::nullopt;
    }
    const size_t rowIndex = static_cast<size_t>(rowId - 1u);
    return rowIndex < rows.size() ? std::optional<size_t>(rowIndex) : std::nullopt;
}

void FolderWindow::SharedDirectoriesDialogState::OnGridSelectionChanged(RedSalamander::DxUi::Grid& sender)
{
    selectedIndex = kSharedDirectoriesNoSelection;
    for (size_t rowIndex = 0u; rowIndex < rows.size(); ++rowIndex)
    {
        if (sender.GetSelectionModel().IsSelected(GetStableRowId(rowIndex)))
        {
            selectedIndex = rowIndex;
            break;
        }
    }
    UpdateEmptyState();
}

void FolderWindow::SharedDirectoriesDialogState::OnGridRowActivated(RedSalamander::DxUi::Grid& /*sender*/, size_t rowIndex)
{
    if (rowIndex < rows.size() && hwnd && IsWindow(hwnd.get()) != FALSE)
    {
        selectedIndex = rowIndex;
        PostMessageW(hwnd.get(), WM_COMMAND, MAKEWPARAM(IDC_SHARED_DIRECTORIES_OPEN, BN_CLICKED), 0);
    }
}

void FolderWindow::SharedDirectoriesDialogState::SetRows(std::vector<SharedDirectoryRow> newRows, HRESULT hr)
{
    rows          = std::move(newRows);
    lastError     = hr;
    selectedIndex = rows.empty() ? kSharedDirectoriesNoSelection : 0u;
    if (grid)
    {
        grid->NotifyDataChanged();
        if (rows.empty())
        {
            grid->GetSelectionModel().Clear();
        }
        else
        {
            grid->GetSelectionModel().SetSingle(GetStableRowId(selectedIndex));
            grid->EnsureRowVisible(selectedIndex);
        }
    }
    UpdateEmptyState();
}

bool FolderWindow::SharedDirectoriesDialogState::SelectRow(size_t rowIndex) noexcept
{
    if (! grid || rowIndex >= rows.size())
    {
        return false;
    }
    selectedIndex = rowIndex;
    grid->GetSelectionModel().SetSingle(GetStableRowId(rowIndex));
    grid->EnsureRowVisible(rowIndex);
    UpdateEmptyState();
    if (hwnd && IsWindow(hwnd.get()) != FALSE)
    {
        InvalidateRect(hwnd.get(), nullptr, FALSE);
    }
    return true;
}

bool FolderWindow::SharedDirectoriesDialogState::OnCreate(HWND createdHwnd) noexcept
{
    if (! dxHost.Attach(createdHwnd))
    {
        return false;
    }
    BuildUi();
    ApplyTheme();
    Layout();
    if (grid && ! rows.empty())
    {
        dxHost.SetFocusControl(grid);
    }
    else if (closeButton)
    {
        dxHost.SetFocusControl(closeButton);
    }
    dxHost.SetDefaultButton(openButton);
    dxHost.SetCancelButton(closeButton);
    return true;
}

void FolderWindow::SharedDirectoriesDialogState::BuildUi()
{
    if (root)
    {
        return;
    }

    using namespace RedSalamander::DxUi;

    columns = {
        {L"share", LoadStringResource(nullptr, IDS_SHARED_DIRECTORIES_COLUMN_NAME), 120.0f, 96.0f, GridColumnKind::Text, false, false},
        {L"localPath", LoadStringResource(nullptr, IDS_SHARED_DIRECTORIES_COLUMN_LOCAL_PATH), 260.0f, 160.0f, GridColumnKind::Text, false, false},
        {L"type", LoadStringResource(nullptr, IDS_SHARED_DIRECTORIES_COLUMN_TYPE), 90.0f, 70.0f, GridColumnKind::Text, false, false},
        {L"remark", LoadStringResource(nullptr, IDS_SHARED_DIRECTORIES_COLUMN_REMARK), 180.0f, 120.0f, GridColumnKind::Text, false, false},
    };

    auto rootOwned = std::make_unique<Panel>();
    root           = rootOwned.get();

    grid = root->AddChild<Grid>();
    grid->SetModel(this);
    grid->SetDelegate(this);
    grid->SetSelectionMode(GridSelectionMode::Single);
    grid->SetEmptyStateText(LoadStringOrFallback(IDS_SHARED_DIRECTORIES_EMPTY, L"No shared directories are available."));
    grid->SetLineClamp(1u);

    emptyLabel = root->AddChild<Label>(LoadStringOrFallback(IDS_SHARED_DIRECTORIES_EMPTY, L"No shared directories are available."));
    emptyLabel->SetAlignment(DWRITE_TEXT_ALIGNMENT_LEADING);
    emptyLabel->SetMultiline(true);

    openButton = root->AddChild<Button>(LoadStringOrFallback(IDS_SHARED_DIRECTORIES_OPEN_PATH, L"Open Path"));
    openButton->SetPrimary(true);
    openButton->SetOnClick([this]
    {
        if (hwnd && IsWindow(hwnd.get()) != FALSE)
        {
            PostMessageW(hwnd.get(), WM_COMMAND, MAKEWPARAM(IDC_SHARED_DIRECTORIES_OPEN, BN_CLICKED), 0);
        }
    });

    manageButton = root->AddChild<Button>(LoadStringOrFallback(IDS_SHARED_DIRECTORIES_MANAGE, L"Manage..."));
    manageButton->SetOnClick([this]
    {
        if (hwnd && IsWindow(hwnd.get()) != FALSE)
        {
            PostMessageW(hwnd.get(), WM_COMMAND, MAKEWPARAM(IDC_SHARED_DIRECTORIES_MANAGE, BN_CLICKED), 0);
        }
    });

    closeButton = root->AddChild<Button>(LoadStringOrFallback(IDS_SHARED_DIRECTORIES_CLOSE, L"Close"));
    closeButton->SetOnClick([this]
    {
        if (hwnd && IsWindow(hwnd.get()) != FALSE)
        {
            PostMessageW(hwnd.get(), WM_COMMAND, MAKEWPARAM(IDCANCEL, BN_CLICKED), 0);
        }
    });

    dxHost.SetRoot(std::move(rootOwned));
    SetRows(std::move(rows), lastError);
}

void FolderWindow::SharedDirectoriesDialogState::ApplyTheme() noexcept
{
    palette = MakeAppThemeDxPalette(theme, theme.windowBackground);
    dxHost.SetTheme(palette);
    if (hwnd && IsWindow(hwnd.get()) != FALSE)
    {
        ApplyWindowChromeTheme(hwnd.get(), theme, WindowBackdropTarget::Tool, GetActiveWindow() == hwnd.get());
        RedrawWindow(hwnd.get(), nullptr, nullptr, RDW_FRAME | RDW_INVALIDATE | RDW_UPDATENOW);
    }
}

void FolderWindow::SharedDirectoriesDialogState::Layout() noexcept
{
    if (! root)
    {
        return;
    }

    const D2D1_RECT_F client = dxHost.GetClientBoundsDip();
    root->SetBounds(client);

    constexpr float kMarginDip       = 20.0f;
    constexpr float kGapDip          = 12.0f;
    constexpr float kButtonWidthDip  = 112.0f;
    constexpr float kButtonHeightDip = 34.0f;

    const float left       = client.left + kMarginDip;
    const float right      = std::max(left, client.right - kMarginDip);
    const float bottom     = std::max(client.top, client.bottom - kMarginDip);
    const float buttonTop  = std::max(client.top + kMarginDip, bottom - kButtonHeightDip);
    const float contentBot = std::max(client.top + kMarginDip, buttonTop - kGapDip);

    if (grid)
    {
        grid->SetBounds(D2D1::RectF(left, client.top + kMarginDip, right, contentBot));
    }
    if (emptyLabel)
    {
        emptyLabel->SetBounds(D2D1::RectF(left, client.top + kMarginDip, right, std::min(contentBot, client.top + kMarginDip + 96.0f)));
    }

    float x = right;
    if (closeButton)
    {
        closeButton->SetBounds(D2D1::RectF(x - kButtonWidthDip, buttonTop, x, buttonTop + kButtonHeightDip));
        x -= kButtonWidthDip + kGapDip;
    }
    if (manageButton)
    {
        manageButton->SetBounds(D2D1::RectF(x - kButtonWidthDip, buttonTop, x, buttonTop + kButtonHeightDip));
        x -= kButtonWidthDip + kGapDip;
    }
    if (openButton)
    {
        openButton->SetBounds(D2D1::RectF(x - kButtonWidthDip, buttonTop, x, buttonTop + kButtonHeightDip));
    }
}

void FolderWindow::SharedDirectoriesDialogState::UpdateEmptyState() noexcept
{
    const bool empty = rows.empty();
    if (grid)
    {
        grid->SetVisible(! empty);
    }
    if (emptyLabel)
    {
        const UINT emptyStringId = lastError == HRESULT_FROM_WIN32(ERROR_ACCESS_DENIED) ? IDS_SHARED_DIRECTORIES_ACCESS_DENIED : IDS_SHARED_DIRECTORIES_EMPTY;
        emptyLabel->SetText(LoadStringOrFallback(emptyStringId,
                                                 lastError == HRESULT_FROM_WIN32(ERROR_ACCESS_DENIED)
                                                     ? L"Shared directories could not be listed because access was denied."
                                                     : L"No shared directories are available."));
        emptyLabel->SetVisible(empty);
    }
    if (openButton)
    {
        const bool canOpen = selectedIndex != kSharedDirectoriesNoSelection && selectedIndex < rows.size() && rows[selectedIndex].openable;
        openButton->SetEnabled(canOpen);
    }
}

void FolderWindow::SharedDirectoriesDialogState::OnNcDestroy(HWND destroyedHwnd) noexcept
{
    destroyed = true;
    dxHost.Detach();
    if (hwnd.get() == destroyedHwnd)
    {
        static_cast<void>(hwnd.release());
    }
    SetWindowLongPtrW(destroyedHwnd, GWLP_USERDATA, 0);
}

#ifdef ENABLE_TESTS
void DebugSetMakeFileListAutomation(const Common::Settings::MakeFileListSettings& options) noexcept
{
    g_makeFileListAutomation = options;
}

void DebugClearMakeFileListAutomation() noexcept
{
    g_makeFileListAutomation.reset();
    g_makeFileListWorkerDelayMs.store(0u, std::memory_order_release);
}

void DebugSetMakeFileListWorkerDelay(uint32_t delayMs) noexcept
{
    g_makeFileListWorkerDelayMs.store(delayMs, std::memory_order_release);
}

bool DebugIsMakeFileListWorkerActive() noexcept
{
    return g_makeFileListWorkerActive.load(std::memory_order_acquire);
}

HWND GetChangeAttributesOptionsPromptHandle() noexcept
{
    const HWND hwnd = g_changeAttributesOptionsPromptWindow.load();
    return hwnd && IsWindow(hwnd) != FALSE ? hwnd : nullptr;
}

bool DebugGetChangeAttributesOptionsPromptSnapshot(ChangeAttributesOptionsPromptDebugSnapshot& out) noexcept
{
    const HWND hwnd    = GetChangeAttributesOptionsPromptHandle();
    const UINT message = GetChangeAttributesOptionsPromptDebugMessage();
    if (! hwnd || message == 0u)
    {
        return false;
    }

    return SendMessageW(hwnd, message, static_cast<WPARAM>(ChangeAttributesOptionsPromptDebugCommand::GetSnapshot), reinterpret_cast<LPARAM>(&out)) != FALSE;
}

bool DebugSetChangeAttributesOptionsPromptState(uint8_t readOnly, uint8_t hidden, uint8_t system, uint8_t archive, bool removeAlternateDataStreams) noexcept
{
    const HWND hwnd    = GetChangeAttributesOptionsPromptHandle();
    const UINT message = GetChangeAttributesOptionsPromptDebugMessage();
    if (! hwnd || message == 0u)
    {
        return false;
    }

    ChangeAttributesOptionsPromptDebugStatePayload payload{};
    payload.readOnly                   = readOnly;
    payload.hidden                     = hidden;
    payload.system                     = system;
    payload.archive                    = archive;
    payload.removeAlternateDataStreams = removeAlternateDataStreams;
    return SendMessageW(hwnd, message, static_cast<WPARAM>(ChangeAttributesOptionsPromptDebugCommand::SetState), reinterpret_cast<LPARAM>(&payload)) != FALSE;
}

bool DebugCycleChangeAttributesOptionsPromptArchive() noexcept
{
    const HWND hwnd    = GetChangeAttributesOptionsPromptHandle();
    const UINT message = GetChangeAttributesOptionsPromptDebugMessage();
    return hwnd && message != 0u && SendMessageW(hwnd, message, static_cast<WPARAM>(ChangeAttributesOptionsPromptDebugCommand::CycleArchive), 0) != FALSE;
}

bool DebugConfirmChangeAttributesOptionsPrompt() noexcept
{
    const HWND hwnd    = GetChangeAttributesOptionsPromptHandle();
    const UINT message = GetChangeAttributesOptionsPromptDebugMessage();
    return PostDxUiPromptCloseDebugCommand(hwnd, message, static_cast<WPARAM>(ChangeAttributesOptionsPromptDebugCommand::Confirm));
}

bool DebugCancelChangeAttributesOptionsPrompt() noexcept
{
    const HWND hwnd    = GetChangeAttributesOptionsPromptHandle();
    const UINT message = GetChangeAttributesOptionsPromptDebugMessage();
    return PostDxUiPromptCloseDebugCommand(hwnd, message, static_cast<WPARAM>(ChangeAttributesOptionsPromptDebugCommand::Cancel));
}

HWND GetMakeFileListOptionsPromptHandle() noexcept
{
    const HWND hwnd = g_makeFileListOptionsPromptWindow.load();
    return hwnd && IsWindow(hwnd) != FALSE ? hwnd : nullptr;
}

bool DebugGetMakeFileListOptionsPromptSnapshot(MakeFileListOptionsPromptDebugSnapshot& out) noexcept
{
    const HWND hwnd    = GetMakeFileListOptionsPromptHandle();
    const UINT message = GetMakeFileListOptionsPromptDebugMessage();
    if (! hwnd || message == 0u)
    {
        return false;
    }

    return SendMessageW(hwnd, message, static_cast<WPARAM>(MakeFileListOptionsPromptDebugCommand::GetSnapshot), reinterpret_cast<LPARAM>(&out)) != FALSE;
}

bool DebugSetMakeFileListOptionsPromptState(uint8_t sourceMode,
                                            bool recursive,
                                            uint8_t format,
                                            uint8_t outputTarget,
                                            std::wstring_view textMacro,
                                            std::wstring_view outputFile,
                                            bool includeName,
                                            bool includeFullPath,
                                            bool includeSize,
                                            bool includeModified,
                                            bool includeAttributes,
                                            bool includeDirectories) noexcept
{
    const HWND hwnd    = GetMakeFileListOptionsPromptHandle();
    const UINT message = GetMakeFileListOptionsPromptDebugMessage();
    if (! hwnd || message == 0u)
    {
        return false;
    }

    MakeFileListOptionsPromptDebugStatePayload payload{};
    payload.sourceMode   = sourceMode;
    payload.recursive    = recursive;
    payload.format       = format;
    payload.outputTarget = outputTarget;
    payload.textMacro.assign(textMacro);
    payload.outputFile.assign(outputFile);
    payload.includeName        = includeName;
    payload.includeFullPath    = includeFullPath;
    payload.includeSize        = includeSize;
    payload.includeModified    = includeModified;
    payload.includeAttributes  = includeAttributes;
    payload.includeDirectories = includeDirectories;
    return SendMessageW(hwnd, message, static_cast<WPARAM>(MakeFileListOptionsPromptDebugCommand::SetState), reinterpret_cast<LPARAM>(&payload)) != FALSE;
}

bool DebugConfirmMakeFileListOptionsPrompt() noexcept
{
    const HWND hwnd    = GetMakeFileListOptionsPromptHandle();
    const UINT message = GetMakeFileListOptionsPromptDebugMessage();
    return PostDxUiPromptCloseDebugCommand(hwnd, message, static_cast<WPARAM>(MakeFileListOptionsPromptDebugCommand::Confirm));
}

bool DebugCancelMakeFileListOptionsPrompt() noexcept
{
    const HWND hwnd    = GetMakeFileListOptionsPromptHandle();
    const UINT message = GetMakeFileListOptionsPromptDebugMessage();
    return PostDxUiPromptCloseDebugCommand(hwnd, message, static_cast<WPARAM>(MakeFileListOptionsPromptDebugCommand::Cancel));
}

HWND GetArchivePackPromptHandle() noexcept
{
    const HWND hwnd = g_archivePackPromptWindow.load();
    return hwnd && IsWindow(hwnd) != FALSE ? hwnd : nullptr;
}

bool DebugGetArchivePackPromptSnapshot(ArchivePackPromptDebugSnapshot& out) noexcept
{
    const HWND hwnd    = GetArchivePackPromptHandle();
    const UINT message = GetArchivePackPromptDebugMessage();
    return hwnd && message != 0u &&
           SendMessageW(hwnd, message, static_cast<WPARAM>(ArchivePackPromptDebugCommand::GetSnapshot), reinterpret_cast<LPARAM>(&out)) != FALSE;
}

bool DebugSetArchivePackPromptPackerIndex(size_t index) noexcept
{
    const HWND hwnd    = GetArchivePackPromptHandle();
    const UINT message = GetArchivePackPromptDebugMessage();
    return hwnd && message != 0u &&
           SendMessageW(hwnd, message, static_cast<WPARAM>(ArchivePackPromptDebugCommand::SetPackerIndex), static_cast<LPARAM>(index)) != FALSE;
}

bool DebugSetArchivePackPromptArchivePath(std::wstring_view path) noexcept
{
    const HWND hwnd    = GetArchivePackPromptHandle();
    const UINT message = GetArchivePackPromptDebugMessage();
    if (! hwnd || message == 0u)
    {
        return false;
    }

    ArchivePackPromptArchivePathPayload payload{};
    payload.archivePath.assign(path);
    return SendMessageW(hwnd, message, static_cast<WPARAM>(ArchivePackPromptDebugCommand::SetArchivePath), reinterpret_cast<LPARAM>(&payload)) != FALSE;
}

bool DebugSetArchivePackPromptDeleteAfter(bool deleteAfterPacking) noexcept
{
    const HWND hwnd    = GetArchivePackPromptHandle();
    const UINT message = GetArchivePackPromptDebugMessage();
    return hwnd && message != 0u &&
           SendMessageW(hwnd, message, static_cast<WPARAM>(ArchivePackPromptDebugCommand::SetDeleteAfter), deleteAfterPacking ? 1 : 0) != FALSE;
}

bool DebugConfirmArchivePackPrompt() noexcept
{
    const HWND hwnd    = GetArchivePackPromptHandle();
    const UINT message = GetArchivePackPromptDebugMessage();
    return PostDxUiPromptCloseDebugCommand(hwnd, message, static_cast<WPARAM>(ArchivePackPromptDebugCommand::Confirm));
}

bool DebugCancelArchivePackPrompt() noexcept
{
    const HWND hwnd    = GetArchivePackPromptHandle();
    const UINT message = GetArchivePackPromptDebugMessage();
    return PostDxUiPromptCloseDebugCommand(hwnd, message, static_cast<WPARAM>(ArchivePackPromptDebugCommand::Cancel));
}

HWND GetArchiveUnpackPromptHandle() noexcept
{
    const HWND hwnd = g_archiveUnpackPromptWindow.load();
    return hwnd && IsWindow(hwnd) != FALSE ? hwnd : nullptr;
}

bool DebugGetArchiveUnpackPromptSnapshot(ArchiveUnpackPromptDebugSnapshot& out) noexcept
{
    const HWND hwnd    = GetArchiveUnpackPromptHandle();
    const UINT message = GetArchiveUnpackPromptDebugMessage();
    return hwnd && message != 0u &&
           SendMessageW(hwnd, message, static_cast<WPARAM>(ArchiveUnpackPromptDebugCommand::GetSnapshot), reinterpret_cast<LPARAM>(&out)) != FALSE;
}

bool DebugSetArchiveUnpackPromptDestinationPath(std::wstring_view path) noexcept
{
    const HWND hwnd    = GetArchiveUnpackPromptHandle();
    const UINT message = GetArchiveUnpackPromptDebugMessage();
    if (! hwnd || message == 0u)
    {
        return false;
    }

    ArchiveUnpackPromptTextPayload payload{};
    payload.text.assign(path);
    return SendMessageW(hwnd, message, static_cast<WPARAM>(ArchiveUnpackPromptDebugCommand::SetDestinationPath), reinterpret_cast<LPARAM>(&payload)) != FALSE;
}

bool DebugSetArchiveUnpackPromptMask(std::wstring_view mask) noexcept
{
    const HWND hwnd    = GetArchiveUnpackPromptHandle();
    const UINT message = GetArchiveUnpackPromptDebugMessage();
    if (! hwnd || message == 0u)
    {
        return false;
    }

    ArchiveUnpackPromptTextPayload payload{};
    payload.text.assign(mask);
    return SendMessageW(hwnd, message, static_cast<WPARAM>(ArchiveUnpackPromptDebugCommand::SetMask), reinterpret_cast<LPARAM>(&payload)) != FALSE;
}

bool DebugSetArchiveUnpackPromptReplaceExisting(bool replaceExisting) noexcept
{
    const HWND hwnd    = GetArchiveUnpackPromptHandle();
    const UINT message = GetArchiveUnpackPromptDebugMessage();
    return hwnd && message != 0u &&
           SendMessageW(hwnd, message, static_cast<WPARAM>(ArchiveUnpackPromptDebugCommand::SetConflictPolicy), replaceExisting ? 1 : 0) != FALSE;
}

bool DebugSetArchiveUnpackPromptDeleteAfter(bool deleteAfterUnpacking) noexcept
{
    const HWND hwnd    = GetArchiveUnpackPromptHandle();
    const UINT message = GetArchiveUnpackPromptDebugMessage();
    return hwnd && message != 0u &&
           SendMessageW(hwnd, message, static_cast<WPARAM>(ArchiveUnpackPromptDebugCommand::SetDeleteAfter), deleteAfterUnpacking ? 1 : 0) != FALSE;
}

bool DebugConfirmArchiveUnpackPrompt() noexcept
{
    const HWND hwnd    = GetArchiveUnpackPromptHandle();
    const UINT message = GetArchiveUnpackPromptDebugMessage();
    return PostDxUiPromptCloseDebugCommand(hwnd, message, static_cast<WPARAM>(ArchiveUnpackPromptDebugCommand::Confirm));
}

bool DebugCancelArchiveUnpackPrompt() noexcept
{
    const HWND hwnd    = GetArchiveUnpackPromptHandle();
    const UINT message = GetArchiveUnpackPromptDebugMessage();
    return PostDxUiPromptCloseDebugCommand(hwnd, message, static_cast<WPARAM>(ArchiveUnpackPromptDebugCommand::Cancel));
}
#endif

void FolderWindow::CommandCreateDirectory(Pane pane)
{
    SetActivePane(pane);
    PaneState& state = pane == Pane::Left ? _leftPane : _rightPane;
    if (! state.fileSystem)
    {
        return;
    }

    HWND ownerWindow = GetOwnerWindowOrSelf(_hWnd.get());
    std::wstring pluginName;
    if (ownerWindow)
    {
        FileSystemPluginManager& pluginManager = FileSystemPluginManager::GetInstance();
        const auto& plugins                    = pluginManager.GetPlugins();
        pluginName                             = TryGetFileSystemPluginDisplayName(plugins, state.pluginId, state.pluginShortId);
    }

    const auto folder = state.folderView.GetFolderPath();
    if (! folder)
    {
        return;
    }

    const std::filesystem::path base = folder.value();

    const HWND folderViewToRestore    = state.hFolderView.get();
    const auto restoreFolderViewFocus = wil::scope_exit([this, pane, folderViewToRestore]() noexcept
    {
        if (folderViewToRestore && IsWindow(folderViewToRestore) != FALSE)
        {
            SetActivePane(pane);
            FocusPaneFolderView(pane);
        }
    });

    wil::com_ptr<IFileSystemDirectoryOperations> dirOps;
    state.fileSystem->QueryInterface(__uuidof(IFileSystemDirectoryOperations), dirOps.put_void());

    const bool canUseWin32 = IsFilePluginShortId(state.pluginShortId) && LooksLikeWindowsAbsolutePath(base.wstring());
    if (! dirOps && ! canUseWin32)
    {
        std::wstring title = LoadStringResource(nullptr, IDS_CAPTION_ERROR);
        std::wstring message;
        if (! pluginName.empty())
        {
            message = FormatStringResource(nullptr, IDS_FMT_PANE_CREATE_DIR_UNSUPPORTED_PLUGIN, pluginName);
        }
        if (message.empty())
        {
            message = LoadStringResource(nullptr, IDS_MSG_PANE_CREATE_DIR_UNSUPPORTED);
        }

        state.folderView.ShowAlertOverlay(FolderView::ErrorOverlayKind::Operation, FolderView::OverlaySeverity::Error, std::move(title), std::move(message));
        return;
    }

    const std::wstring defaultNameBase = LoadStringResource(nullptr, IDS_NEW_FOLDER_DEFAULT_NAME);
    if (defaultNameBase.empty())
    {
        return;
    }

    const bool treatSuggestedNamesCaseInsensitive = IsFilePluginShortId(state.pluginShortId);
    const std::wstring initialName = ResolveInitialCreateDirectoryName(state.fileSystem, base, defaultNameBase, treatSuggestedNamesCaseInsensitive);

    if (! ownerWindow)
    {
        ownerWindow = _hWnd.get();
    }

    const std::filesystem::path displayPath = NavigationLocation::FormatHistoryPath(state.pluginShortId, state.instanceContext, base);
    const auto folderName                   = PromptForCreateDirectoryName(ownerWindow, displayPath.wstring(), initialName, _theme);
    if (! folderName.has_value())
    {
        return;
    }

    const std::wstring requestedName = folderName.value();
    unsigned int suggestedSuffix     = 0u;
    const bool autoSuffix            = TryParseCreateDirectorySuggestedSuffix(requestedName, defaultNameBase, suggestedSuffix);

    const int maxAttempts = autoSuffix ? 1000 : 1;
    for (int attempt = 0; attempt < maxAttempts; ++attempt)
    {
        std::wstring candidateName = requestedName;
        if (autoSuffix)
        {
            candidateName = BuildCreateDirectorySuggestedName(defaultNameBase, suggestedSuffix + static_cast<unsigned int>(attempt));
        }

        const std::filesystem::path newFolderPath = base / std::filesystem::path(candidateName);
        if (newFolderPath.empty())
        {
            continue;
        }

        HRESULT hr = S_OK;
        if (dirOps)
        {
            hr = dirOps->CreateDirectory(newFolderPath.c_str());
        }
        else
        {
            if (::CreateDirectoryW(newFolderPath.c_str(), nullptr) == 0)
            {
                const DWORD error = GetLastError();
                hr                = HRESULT_FROM_WIN32(error);
            }
        }

        if (SUCCEEDED(hr))
        {
            const std::wstring focusName = newFolderPath.filename().wstring();
            if (! focusName.empty())
            {
                state.folderView.RememberFocusedItemForFolder(base, focusName);
            }

            DirectoryInfoCache& cache = DirectoryInfoCache::GetInstance();
            cache.NotifyPathCreated(state.fileSystem.get(), newFolderPath);
            state.folderView.ForceRefresh();

            const Pane otherPane   = pane == Pane::Left ? Pane::Right : Pane::Left;
            PaneState& otherState  = otherPane == Pane::Left ? _leftPane : _rightPane;
            const auto otherFolder = otherState.folderView.GetFolderPath();
            if (otherState.fileSystem && otherFolder.has_value() && OrdinalString::EqualsNoCasePath(otherFolder.value(), base) &&
                EqualsNoCase(otherState.pluginId, state.pluginId) && EqualsNoCase(otherState.instanceContext, state.instanceContext) &&
                ! cache.IsFolderWatched(otherState.fileSystem.get(), base))
            {
                otherState.folderView.ForceRefresh();
            }
            return;
        }

        if (hr == E_NOTIMPL)
        {
            std::wstring title = LoadStringResource(nullptr, IDS_CAPTION_ERROR);
            std::wstring message;
            if (! pluginName.empty())
            {
                message = FormatStringResource(nullptr, IDS_FMT_PANE_CREATE_DIR_UNSUPPORTED_PLUGIN, pluginName);
            }
            if (message.empty())
            {
                message = LoadStringResource(nullptr, IDS_MSG_PANE_CREATE_DIR_UNSUPPORTED);
            }

            state.folderView.ShowAlertOverlay(
                FolderView::ErrorOverlayKind::Operation, FolderView::OverlaySeverity::Error, std::move(title), std::move(message));
            return;
        }

        constexpr HRESULT alreadyExistsHr = HRESULT_FROM_WIN32(ERROR_ALREADY_EXISTS);
        constexpr HRESULT fileExistsHr    = HRESULT_FROM_WIN32(ERROR_FILE_EXISTS);
        if (autoSuffix && (hr == alreadyExistsHr || hr == fileExistsHr))
        {
            continue;
        }

        std::wstring title   = LoadStringResource(nullptr, IDS_CAPTION_ERROR);
        std::wstring message = FormatStringResource(nullptr, IDS_FMT_PANE_CREATE_DIR_FAILED, newFolderPath.wstring(), static_cast<unsigned long>(hr));
        state.folderView.ShowAlertOverlay(
            FolderView::ErrorOverlayKind::Operation, FolderView::OverlaySeverity::Error, std::move(title), std::move(message), hr);
        return;
    }
}

void FolderWindow::CommandEditNew(Pane pane)
{
    SetActivePane(pane);
    PaneState& state = pane == Pane::Left ? _leftPane : _rightPane;
    if (! state.fileSystem)
    {
        return;
    }

    HWND ownerWindow = GetOwnerWindowOrSelf(_hWnd.get());
    std::wstring pluginName;
    if (ownerWindow)
    {
        FileSystemPluginManager& pluginManager = FileSystemPluginManager::GetInstance();
        const auto& plugins                    = pluginManager.GetPlugins();
        pluginName                             = TryGetFileSystemPluginDisplayName(plugins, state.pluginId, state.pluginShortId);
    }

    const auto folder = state.folderView.GetFolderPath();
    if (! folder)
    {
        return;
    }

    const std::filesystem::path base = folder.value();

    wil::com_ptr<IFileSystemIO> fileIo;
    state.fileSystem->QueryInterface(__uuidof(IFileSystemIO), fileIo.put_void());
    if (! fileIo)
    {
        std::wstring title = LoadStringResource(nullptr, IDS_CAPTION_ERROR);
        std::wstring message;
        if (! pluginName.empty())
        {
            message = FormatStringResource(nullptr, IDS_FMT_PANE_EDIT_NEW_UNSUPPORTED_PLUGIN, pluginName);
        }
        if (message.empty())
        {
            message = LoadStringResource(nullptr, IDS_MSG_PANE_EDIT_NEW_UNSUPPORTED);
        }

        state.folderView.ShowAlertOverlay(FolderView::ErrorOverlayKind::Operation, FolderView::OverlaySeverity::Error, std::move(title), std::move(message));
        return;
    }

    if (! ownerWindow)
    {
        ownerWindow = _hWnd.get();
    }

    const std::filesystem::path displayPath = NavigationLocation::FormatHistoryPath(state.pluginShortId, state.instanceContext, base);
    const std::wstring computerName         = GetComputerNameTextForFileActions();
    const auto promptResult =
        PromptForEditNewFile(ownerWindow, base, displayPath.wstring(), _settings ? &_settings->fileActions.editors : nullptr, computerName, _theme);
    if (! promptResult.has_value())
    {
        return;
    }

    const std::wstring requestedName            = promptResult->fileName;
    const std::optional<UINT> validationMessage = ResolveEditNewValidationMessageId(requestedName, base);
    if (validationMessage.has_value())
    {
        std::wstring title   = LoadStringResource(nullptr, IDS_CAPTION_ERROR);
        std::wstring message = LoadStringResource(nullptr, validationMessage.value());
        state.folderView.ShowAlertOverlay(FolderView::ErrorOverlayKind::Operation, FolderView::OverlaySeverity::Error, std::move(title), std::move(message));
        return;
    }

    const std::filesystem::path newFilePath = base / std::filesystem::path(requestedName);
    const auto createStartedAt              = std::chrono::steady_clock::now();

    wil::com_ptr<IFileWriter> writer;
    HRESULT hr = fileIo->CreateFileWriter(newFilePath.c_str(), FILESYSTEM_FLAG_NONE, writer.put());
    if (SUCCEEDED(hr) && writer)
    {
        hr = writer->Commit();
    }
    Debug::Perf::Emit(L"fileaction.editnew.create_file_us", L"", Debug::Perf::ElapsedUs(createStartedAt), 1u, 0u, hr);

    if (FAILED(hr))
    {
        std::wstring title = LoadStringResource(nullptr, IDS_CAPTION_ERROR);
        std::wstring message;
        if (hr == E_NOTIMPL)
        {
            if (! pluginName.empty())
            {
                message = FormatStringResource(nullptr, IDS_FMT_PANE_EDIT_NEW_UNSUPPORTED_PLUGIN, pluginName);
            }
            if (message.empty())
            {
                message = LoadStringResource(nullptr, IDS_MSG_PANE_EDIT_NEW_UNSUPPORTED);
            }
        }
        else
        {
            message = FormatStringResource(nullptr, IDS_FMT_PANE_EDIT_NEW_FAILED, newFilePath.wstring(), static_cast<unsigned long>(hr));
        }

        state.folderView.ShowAlertOverlay(
            FolderView::ErrorOverlayKind::Operation, FolderView::OverlaySeverity::Error, std::move(title), std::move(message), hr);
        return;
    }

    const std::wstring focusName = newFilePath.filename().wstring();
    if (! focusName.empty())
    {
        state.folderView.RememberFocusedItemForFolder(base, focusName);
    }

    DirectoryInfoCache& cache = DirectoryInfoCache::GetInstance();
    cache.NotifyPathCreated(state.fileSystem.get(), newFilePath);
    state.folderView.ForceRefresh();

    const Pane otherPane   = pane == Pane::Left ? Pane::Right : Pane::Left;
    PaneState& otherState  = otherPane == Pane::Left ? _leftPane : _rightPane;
    const auto otherFolder = otherState.folderView.GetFolderPath();
    if (otherState.fileSystem && otherFolder.has_value() && OrdinalString::EqualsNoCasePath(otherFolder.value(), base) &&
        EqualsNoCase(otherState.pluginId, state.pluginId) && EqualsNoCase(otherState.instanceContext, state.instanceContext) &&
        ! cache.IsFolderWatched(otherState.fileSystem.get(), base))
    {
        otherState.folderView.ForceRefresh();
    }

    if (! promptResult->editorActionId.empty())
    {
        std::vector<std::filesystem::path> selectedPaths{newFilePath};
        if (! TryEditFileWithEditor(pane, newFilePath, selectedPaths, promptResult->editorActionId, false))
        {
            static_cast<void>(ShowRecordedFileActionFailureOverlay(pane));
        }
    }
}

std::vector<FolderWindow::ShellNewTemplateMenuItem> FolderWindow::CollectShellNewTemplateMenuItems(Pane pane) const
{
    const PaneState& state = pane == Pane::Left ? _leftPane : _rightPane;
    if (! IsFilePluginShortId(state.pluginShortId))
    {
        return {};
    }

    std::vector<ShellNewTemplateDefinition> templates;
#ifdef ENABLE_TESTS
    if (_debugShellNewTemplates.has_value())
    {
        templates = _debugShellNewTemplates.value();
    }
    else
#endif
    {
        templates = EnumerateShellNewTemplatesFromRegistry();
    }

    std::vector<ShellNewTemplateMenuItem> items;
    items.reserve(templates.size());
    for (const ShellNewTemplateDefinition& entry : templates)
    {
        if (! entry.id.empty() && ! entry.displayName.empty())
        {
            items.push_back(ShellNewTemplateMenuItem{.id = entry.id, .displayName = entry.displayName});
        }
    }
    return items;
}

void FolderWindow::CommandNewFromShellTemplate(Pane pane, std::wstring_view templateId)
{
    SetActivePane(pane);
    PaneState& state = pane == Pane::Left ? _leftPane : _rightPane;

    HWND ownerWindow       = GetOwnerWindowOrSelf(_hWnd.get());
    const auto showWarning = [&](std::wstring message, HRESULT hr = S_OK) noexcept
    {
        Debug::Perf::Scope perf(L"shellnew.feedback_us");
        std::wstring title = LoadStringResource(nullptr, IDS_SHELL_ACTION_UNAVAILABLE_TITLE);
        state.folderView.ShowAlertOverlay(
            FolderView::ErrorOverlayKind::Operation, FolderView::OverlaySeverity::Warning, std::move(title), std::move(message), hr);
    };

    const auto folder = state.folderView.GetFolderPath();
    if (! state.fileSystem || ! folder.has_value() || ! IsFilePluginShortId(state.pluginShortId))
    {
        showWarning(LoadStringResource(nullptr, IDS_MSG_SHELL_NEW_LOCAL_FOLDER_REQUIRED));
        return;
    }

    std::vector<ShellNewTemplateDefinition> templates;
#ifdef ENABLE_TESTS
    if (_debugShellNewTemplates.has_value())
    {
        templates = _debugShellNewTemplates.value();
    }
    else
#endif
    {
        templates = EnumerateShellNewTemplatesFromRegistry();
    }

    if (templates.empty())
    {
        showWarning(LoadStringResource(nullptr, IDS_MSG_SHELL_NEW_NO_TEMPLATES));
        return;
    }

    const std::wstring requestedTemplateId(templateId);
    const auto templateIt = std::find_if(templates.begin(), templates.end(), [&](const ShellNewTemplateDefinition& entry) noexcept {
        return ! requestedTemplateId.empty() && EqualsNoCase(entry.id, requestedTemplateId);
    });
    if (templateIt == templates.end())
    {
        const std::wstring unavailableId = requestedTemplateId.empty() ? std::wstring(L"<none>") : requestedTemplateId;
        showWarning(FormatStringResource(nullptr, IDS_FMT_SHELL_NEW_TEMPLATE_UNAVAILABLE, unavailableId));
        return;
    }

    wil::com_ptr<IFileSystemIO> fileIo;
    state.fileSystem->QueryInterface(__uuidof(IFileSystemIO), fileIo.put_void());
    if (! fileIo)
    {
        showWarning(LoadStringResource(nullptr, IDS_MSG_SHELL_NEW_LOCAL_FOLDER_REQUIRED), E_NOTIMPL);
        return;
    }

    const std::filesystem::path base = folder.value();
    std::optional<EditNewPromptResult> promptResult;
#ifdef ENABLE_TESTS
    if (_debugNextShellNewFileName.has_value())
    {
        promptResult = EditNewPromptResult{.fileName = _debugNextShellNewFileName.value()};
        _debugNextShellNewFileName.reset();
    }
    else
#endif
    {
        if (! ownerWindow)
        {
            ownerWindow = _hWnd.get();
        }
        const std::filesystem::path displayPath = NavigationLocation::FormatHistoryPath(state.pluginShortId, state.instanceContext, base);
        const std::wstring caption = templateIt->displayName.empty() ? LoadStringResource(nullptr, IDS_CMD_NEW_FROM_TEMPLATE) : templateIt->displayName;
        promptResult = PromptForEditNewFile(ownerWindow, base, displayPath.wstring(), nullptr, {}, _theme, templateIt->defaultFileName, caption, false);
    }
    if (! promptResult.has_value())
    {
        return;
    }

    std::wstring requestedName                  = EnsureShellNewExtension(promptResult->fileName, templateIt->extension);
    const std::optional<UINT> validationMessage = ResolveEditNewValidationMessageId(requestedName, base);
    if (validationMessage.has_value())
    {
        showWarning(LoadStringResource(nullptr, validationMessage.value()));
        return;
    }

    const std::filesystem::path newFilePath = base / std::filesystem::path(requestedName);
    const auto createStartedAt              = std::chrono::steady_clock::now();
    uint64_t bytesWritten                   = 0u;

    wil::com_ptr<IFileWriter> writer;
    HRESULT hr = fileIo->CreateFileWriter(newFilePath.c_str(), FILESYSTEM_FLAG_NONE, writer.put());
    if (SUCCEEDED(hr) && writer)
    {
        switch (templateIt->kind)
        {
            case ShellNewTemplateKind::NullFile: break;
            case ShellNewTemplateKind::Data:
                bytesWritten = static_cast<uint64_t>(templateIt->data.size());
                hr           = WriteShellNewTemplateBytes(*writer, templateIt->data.data(), templateIt->data.size());
                break;
            case ShellNewTemplateKind::FileName: hr = WriteShellNewTemplateFile(*writer, templateIt->templateFilePath, bytesWritten); break;
        }
    }
    if (SUCCEEDED(hr) && writer)
    {
        hr = writer->Commit();
    }

    Debug::Perf::Emit(L"shellnew.create_us", ShellNewTemplateKindDetail(templateIt->kind), Debug::Perf::ElapsedUs(createStartedAt), bytesWritten, 1u, hr);

    if (FAILED(hr))
    {
        showWarning(FormatStringResource(nullptr, IDS_FMT_SHELL_NEW_CREATE_FAILED, newFilePath.wstring(), static_cast<unsigned long>(hr)), hr);
        return;
    }

    const std::wstring focusName = newFilePath.filename().wstring();
    if (! focusName.empty())
    {
        state.folderView.RememberFocusedItemForFolder(base, focusName);
    }

    DirectoryInfoCache& cache = DirectoryInfoCache::GetInstance();
    cache.NotifyPathCreated(state.fileSystem.get(), newFilePath);
    state.folderView.ForceRefresh();

    const Pane otherPane   = pane == Pane::Left ? Pane::Right : Pane::Left;
    PaneState& otherState  = otherPane == Pane::Left ? _leftPane : _rightPane;
    const auto otherFolder = otherState.folderView.GetFolderPath();
    if (otherState.fileSystem && otherFolder.has_value() && OrdinalString::EqualsNoCasePath(otherFolder.value(), base) &&
        EqualsNoCase(otherState.pluginId, state.pluginId) && EqualsNoCase(otherState.instanceContext, state.instanceContext) &&
        ! cache.IsFolderWatched(otherState.fileSystem.get(), base))
    {
        otherState.folderView.ForceRefresh();
    }
}

void FolderWindow::CommandRefresh(Pane pane)
{
    SetActivePane(pane);
    PaneState& state = pane == Pane::Left ? _leftPane : _rightPane;
    state.folderView.ForceRefresh();
}

#ifdef ENABLE_TESTS
void FolderWindow::DebugSetShellActionCallback(DebugShellActionCallback callback)
{
    _debugShellActionCallback = std::move(callback);
}

void FolderWindow::DebugSetNextChangeAttributesOptions(std::optional<ChangeAttributesOptions> options)
{
    _debugNextChangeAttributesOptions = std::move(options);
}

std::optional<FolderWindow::ChangeAttributesReport> FolderWindow::DebugGetLastChangeAttributesReport() const
{
    return _debugLastChangeAttributesReport;
}

void FolderWindow::DebugSetShellNewTemplatesForTest(std::optional<std::vector<ShellNewTemplateDefinition>> templates)
{
    _debugShellNewTemplates = std::move(templates);
}

void FolderWindow::DebugSetNextShellNewFileNameForTest(std::optional<std::wstring> fileName)
{
    _debugNextShellNewFileName = std::move(fileName);
}

uint64_t FolderWindow::DebugGetForceRefreshCount(Pane pane) const noexcept
{
    const PaneState& state = pane == Pane::Left ? _leftPane : _rightPane;
    return state.folderView.DebugGetForceRefreshCount();
}

std::wstring_view FolderWindow::DebugGetFocusedItemDisplayName(Pane pane) const noexcept
{
    const PaneState& state = pane == Pane::Left ? _leftPane : _rightPane;
    return state.folderView.DebugGetFocusedDisplayName();
}

bool FolderWindow::DebugGetPaneViewOptionsSnapshot(Pane pane, PaneViewOptionsDebugSnapshot& out) const
{
    const PaneState& state                   = pane == Pane::Left ? _leftPane : _rightPane;
    const FolderView::NameFilterState filter = state.folderView.GetNameFilterState();

    out.fileExtensionsVisible                           = state.folderView.GetFileExtensionsVisible();
    out.navigationBarVisible                            = state.navigationBarVisible;
    out.navigationViewWindowVisible                     = state.hNavigationView && IsWindowVisible(state.hNavigationView.get()) != FALSE;
    out.filterBarVisible                                = state.filterBarVisible;
    out.filterBarWindowVisible                          = state.hFilterBar && IsWindowVisible(state.hFilterBar.get()) != FALSE;
    out.filterEnabled                                   = filter.enabled;
    const FolderView::ThumbnailDebugSnapshot thumbnails = state.folderView.DebugGetThumbnailSnapshot();
    out.thumbnailsVisible                               = thumbnails.visible;
    out.thumbnailTargetDip                              = thumbnails.targetDip;
    out.thumbnailQueuedCount                            = thumbnails.queuedCount;
    out.thumbnailCompletedCount                         = thumbnails.completedCount;
    out.thumbnailFallbackCount                          = thumbnails.fallbackCount;
    out.thumbnailStaleDropCount                         = thumbnails.staleDropCount;
    out.thumbnailPendingCount                           = thumbnails.pendingCount;
    out.thumbnailCacheHitCount                          = thumbnails.cacheHitCount;
    out.thumbnailShellSuccessCount                      = thumbnails.shellSuccessCount;
    out.thumbnailShellCacheHitCount                     = thumbnails.shellCacheHitCount;
    out.thumbnailShellCacheMissCount                    = thumbnails.shellCacheMissCount;
    out.thumbnailShellProviderAllowedCount              = thumbnails.shellProviderAllowedCount;
    out.thumbnailShellProviderTimeoutCount              = thumbnails.shellProviderTimeoutCount;
    out.thumbnailWicSuccessCount                        = thumbnails.wicSuccessCount;
    out.thumbnailWicFactoryCreateCount                  = thumbnails.wicFactoryCreateCount;
    out.thumbnailDecodeFailureCount                     = thumbnails.decodeFailureCount;
    out.thumbnailVisibleApplyCount                      = thumbnails.visibleApplyCount;
    out.thumbnailVisibleItemCount                       = thumbnails.visibleItemCount;
    out.thumbnailVisibleThumbnailCount                  = thumbnails.visibleThumbnailCount;
    out.thumbnailTotalThumbnailCount                    = thumbnails.totalThumbnailCount;
    out.thumbnailCacheBytes                             = thumbnails.cacheBytes;
    out.thumbnailCacheEvictedCount                      = thumbnails.cacheEvictedCount;
    out.thumbnailCancelCount                            = thumbnails.cancelCount;
    out.thumbnailLastDrawSawThumbnail                   = thumbnails.lastDrawSawThumbnail;
    out.thumbnailLastDrawSourceWidthPx                  = thumbnails.lastDrawSourceWidthPx;
    out.thumbnailLastDrawSourceHeightPx                 = thumbnails.lastDrawSourceHeightPx;
    out.thumbnailLastDrawSlotRectDip                    = thumbnails.lastDrawSlotRectDip;
    out.thumbnailLastDrawRectDip                        = thumbnails.lastDrawRectDip;
    out.iconLastDrawSawIcon                             = thumbnails.lastIconDrawSawIcon;
    out.iconLastDrawSourceWidthPx                       = thumbnails.lastIconDrawSourceWidthPx;
    out.iconLastDrawSourceHeightPx                      = thumbnails.lastIconDrawSourceHeightPx;
    out.iconLastDrawSlotRectDip                         = thumbnails.lastIconDrawSlotRectDip;
    out.iconLastDrawRectDip                             = thumbnails.lastIconDrawRectDip;
    out.filterText                                      = filter.text;
    out.filterBarText                                   = state.filterBarText;
    out.filterBarUsesDxUiHost                           = state.filterBarHost.GetRoot() != nullptr;
    out.filterBarLabelVisible                           = state.filterBarLabel && state.filterBarLabel->IsVisible();
    out.filterBarComboVisible                           = state.filterBarCombo && state.filterBarCombo->IsVisible();
    out.filterBarToggleVisible                          = state.filterBarToggle && state.filterBarToggle->IsVisible();
    out.filterBarToggleChecked                          = state.filterBarToggle && state.filterBarToggle->IsChecked();
    out.filterBarFieldText                              = state.filterBarCombo ? std::wstring(state.filterBarCombo->GetText()) : std::wstring{};
    out.filterBarHistoryItems.clear();
    if (state.filterBarCombo)
    {
        const auto historyItems       = state.filterBarCombo->GetItems();
        out.filterBarHistoryItemCount = historyItems.size();
        out.filterBarHistoryItems.reserve(historyItems.size());
        for (const RedSalamander::DxUi::ComboBox::Item& item : historyItems)
        {
            out.filterBarHistoryItems.push_back(item.value);
        }
    }
    else
    {
        out.filterBarHistoryItemCount = 0u;
    }
    out.focusedItemRealDisplayName   = std::wstring(state.folderView.DebugGetFocusedDisplayName());
    out.focusedItemVisualDisplayName = state.folderView.DebugGetFocusedVisualDisplayName();
    return true;
}

void FolderWindow::DebugSetThumbnailProviderMode(Pane pane, FolderView::DebugThumbnailProviderMode mode) noexcept
{
    PaneState& state = pane == Pane::Left ? _leftPane : _rightPane;
    state.folderView.DebugSetThumbnailProviderMode(mode);
}

bool FolderWindow::DebugSeedThumbnailPendingAndPostThumbnailBitmapMessagesForTest(
    Pane pane, uint64_t pendingCount, uint64_t staleBatchMessageCount, uint64_t staleGenerationMessageCount, uint64_t unaccountedCurrentMessageCount)
{
    PaneState& state = pane == Pane::Left ? _leftPane : _rightPane;
    return state.folderView.DebugSeedThumbnailPendingAndPostThumbnailBitmapMessagesForTest(
        pendingCount, staleBatchMessageCount, staleGenerationMessageCount, unaccountedCurrentMessageCount);
}

bool FolderWindow::DebugGetPreviewPaneSnapshot(PreviewPaneDebugSnapshot& out) const noexcept
{
    out                 = {};
    out.clientRect      = {0, 0, _clientSize.cx, _clientSize.cy};
    out.functionBarRect = _functionBarRect;

    const Pane sourcePane        = _previewSourcePane.value_or(_activePane);
    const Pane hostPane          = OppositePane(sourcePane);
    const PaneState& sourceState = sourcePane == Pane::Left ? _leftPane : _rightPane;
    const PaneState& hostState   = hostPane == Pane::Left ? _leftPane : _rightPane;

    out.active                           = _previewSourcePane.has_value();
    out.sourcePane                       = sourcePane;
    out.hostPane                         = hostPane;
    out.tabsVisible                      = hostState.hPreviewTabs && IsWindowVisible(hostState.hPreviewTabs.get()) != FALSE;
    out.tabsUseDxUiHost                  = hostState.previewTabsHost.GetRoot() != nullptr;
    out.previewTabSelected               = hostState.previewTabSelected;
    out.folderTabSelected                = ! hostState.previewTabSelected;
    out.previewContentVisible            = hostState.hPreviewContent && IsWindowVisible(hostState.hPreviewContent.get()) != FALSE;
    out.previewContentUsesDxUiHost       = hostState.previewContentHost.GetRoot() != nullptr;
    out.previewUsesEmbeddedViewer        = hostState.previewViewerInstance != nullptr;
    out.previewPropertiesCardMode        = hostState.previewPropertiesCardMode;
    out.previewPropertiesUsesScrollPanel = hostState.previewPropertiesScroll != nullptr;
    out.previewPropertiesUsesRainbow     = hostState.previewPropertiesUsesRainbow;
    out.previewPropertiesSectionCount    = hostState.previewPropertiesSections.size();
    out.previewPropertiesFieldCount      = hostState.previewPropertiesFieldCount;
    if (hostState.previewPropertiesScroll)
    {
        const D2D1_RECT_F scrollBounds = hostState.previewPropertiesScroll->GetBounds();
        const float viewportH          = (std::max)(0.0f, scrollBounds.bottom - scrollBounds.top);
        const float scrollMaxDip       = (std::max)(0.0f, hostState.previewPropertiesScroll->GetContentHeight() - viewportH);
        out.previewPropertiesCanScroll = hostState.previewPropertiesScroll->NeedsScrollbar();
        out.previewPropertiesScrollOffsetPx =
            static_cast<int>(std::lround(hostState.previewContentHost.DipsToPixels(hostState.previewPropertiesScroll->GetScrollOffset())));
        out.previewPropertiesScrollMaxPx = static_cast<int>(std::lround(hostState.previewContentHost.DipsToPixels(scrollMaxDip)));
    }
    out.folderViewVisible   = hostState.hFolderView && IsWindowVisible(hostState.hFolderView.get()) != FALSE;
    out.previewTabsHwnd     = hostState.hPreviewTabs.get();
    out.previewContentHwnd  = hostState.hPreviewContent.get();
    const HWND embeddedHwnd = hostState.previewViewerInstance ? hostState.previewViewerInstance->embeddedHwnd : nullptr;
    if (IsOwnedPreviewEmbeddedHwnd(hostState, hostState.previewViewerInstance, embeddedHwnd))
    {
        out.previewEmbeddedViewerHwnd = embeddedHwnd;
    }
    if (hostState.hPreviewContent)
    {
        for (HWND child = GetWindow(hostState.hPreviewContent.get(), GW_CHILD); child != nullptr; child = GetWindow(child, GW_HWNDNEXT))
        {
            if (GetParent(child) != hostState.hPreviewContent.get())
            {
                continue;
            }
            ++out.previewDirectChildCount;
            if (IsWindowVisible(child) != FALSE)
            {
                ++out.previewVisibleDirectChildCount;
            }
            if ((GetWindowLongPtrW(child, GWL_STYLE) & WS_VISIBLE) != 0)
            {
                ++out.previewOwnVisibleDirectChildCount;
            }
        }
    }
    out.previewLastOpenCreatedChildCount        = _debugPreviewLastOpenCreatedChildCount;
    out.previewLastOpenDetectedChildCount       = _debugPreviewLastOpenDetectedChildCount;
    out.previewLastOpenHiddenRejectedChildCount = _debugPreviewLastOpenHiddenRejectedChildCount;
    out.previewLastOpenRejectedChildCardinality = _debugPreviewLastOpenRejectedChildCardinality;
    out.previewLastReopenRejectedChildSet       = _debugPreviewLastReopenRejectedChildSet;
    out.previewViewerInstanceId                 = reinterpret_cast<uintptr_t>(hostState.previewViewerInstance);
    out.tabRect                                 = hostPane == Pane::Left ? _leftPreviewTabsRect : _rightPreviewTabsRect;
    out.contentRect                             = hostPane == Pane::Left ? _leftPreviewContentRect : _rightPreviewContentRect;
    out.previewedPath                           = hostState.previewedPath;
    out.previewText                             = hostState.previewText;
#if defined(ENABLE_TESTS) && defined(_DEBUG)
    if (out.previewEmbeddedViewerHwnd && hostState.previewViewerPluginId == L"builtin/viewer-text")
    {
        WndMsg::ViewerTextDebugSnapshot textSnapshot{};
        if (SendMessageW(out.previewEmbeddedViewerHwnd, WndMsg::kViewerTextDebugGetSnapshot, 0u, reinterpret_cast<LPARAM>(&textSnapshot)) != FALSE)
        {
            out.previewText = textSnapshot.textPreview;
        }
    }
#endif
#if defined(ENABLE_TESTS)
    if (hostState.previewTabsControl)
    {
        const auto toClientRect = [&hostState](const D2D1_RECT_F& rect) noexcept
        {
            return RECT{static_cast<LONG>(std::lround(hostState.previewTabsHost.DipsToPixels(rect.left))),
                        static_cast<LONG>(std::lround(hostState.previewTabsHost.DipsToPixels(rect.top))),
                        static_cast<LONG>(std::lround(hostState.previewTabsHost.DipsToPixels(rect.right))),
                        static_cast<LONG>(std::lround(hostState.previewTabsHost.DipsToPixels(rect.bottom)))};
        };
        out.folderTabClientRect         = toClientRect(hostState.previewTabsControl->DebugGetTabRect(0u));
        out.previewTabClientRect        = toClientRect(hostState.previewTabsControl->DebugGetTabRect(1u));
        out.previewCloseClientRect      = toClientRect(hostState.previewTabsControl->DebugGetCloseButtonRect(1u));
        out.previewCloseButtonVisible   = hostState.previewTabsControl->DebugIsCloseButtonVisible(1u);
        out.previewTabsHasHeaderDivider = hostState.previewTabsControl->DebugHasHeaderDivider();
    }
#endif
    if (hostState.previewTabsHost.HasTooltip())
    {
        out.previewTabsTooltipText = std::wstring(hostState.previewTabsHost.GetTooltipText());
    }
#if defined(ENABLE_TESTS)
    out.previewTabsPendingTooltipText = std::wstring(hostState.previewTabsHost.DebugGetPendingTooltipText());
#endif
    out.previewViewerPluginId    = hostState.previewViewerPluginId;
    out.previewBytes             = hostState.previewBytes;
    out.sourceFocusedDisplayName = std::wstring(sourceState.folderView.DebugGetFocusedDisplayName());
    return true;
}

bool FolderWindow::DebugSetPreviewPaneTab(Pane hostPane, bool previewTab) noexcept
{
    if (! _previewSourcePane.has_value() || OppositePane(_previewSourcePane.value()) != hostPane)
    {
        return false;
    }

    const PaneState& host = hostPane == Pane::Left ? _leftPane : _rightPane;
    if (! host.previewTabsVisible)
    {
        return false;
    }

    SetPreviewPaneTab(hostPane, previewTab);
    return true;
}

bool FolderWindow::DebugScrollPreviewPropertiesByWheelDetents(Pane hostPane, int detents) noexcept
{
    if (detents == 0)
    {
        return false;
    }

    if (! _previewSourcePane.has_value() || OppositePane(_previewSourcePane.value()) != hostPane)
    {
        return false;
    }

    PaneState& host = hostPane == Pane::Left ? _leftPane : _rightPane;
    if (! host.previewPropertiesScroll || ! host.previewPropertiesScroll->NeedsScrollbar())
    {
        return false;
    }

    constexpr float kWheelStepDip = 48.0f;
    const float before            = host.previewPropertiesScroll->GetScrollOffset();
    host.previewPropertiesScroll->SetScrollOffset(before - (static_cast<float>(detents) * kWheelStepDip));
    host.previewContentHost.Invalidate();

    FocusPaneFolderView(_previewSourcePane.value());
    return host.previewPropertiesScroll->GetScrollOffset() != before;
}

bool FolderWindow::DebugAdvancePreviewTabsTooltipDelayForTest(Pane hostPane) noexcept
{
    PaneState& host = hostPane == Pane::Left ? _leftPane : _rightPane;
    if (! host.previewTabsControl || ! host.hPreviewTabs || IsWindow(host.hPreviewTabs.get()) == FALSE)
    {
        return false;
    }

    return host.previewTabsHost.DebugAdvanceTooltipDelayForTest();
}

bool FolderWindow::DebugGetOpenedFilesDialogSnapshot(OpenedFilesDebugSnapshot& out) const noexcept
{
    out               = {};
    out.selectedIndex = kOpenedFilesNoSelection;

    if (! _openedFilesDialog || ! _openedFilesDialog->hwnd || IsWindow(_openedFilesDialog->hwnd.get()) == FALSE)
    {
        return false;
    }

    const HWND dialog                  = _openedFilesDialog->hwnd.get();
    out.visible                        = IsWindowVisible(dialog) != FALSE;
    out.usesDxUiHost                   = _openedFilesDialog->dxHost.GetRoot() != nullptr;
    out.visibleChildWindowCount        = CountVisibleChildWindowsLocal(dialog);
    out.visibleNativeChildControlCount = CountVisibleNativeChildControlWindowsLocal(dialog);
    out.dialogClassName                = GetWindowClassNameLocal(dialog);
    out.themeWindowBackground          = _openedFilesDialog->theme.windowBackground;
    out.themeText                      = _openedFilesDialog->theme.menu.text;
    out.emptyStateVisible              = _openedFilesDialog->emptyLabel && _openedFilesDialog->emptyLabel->IsVisible();
    out.selectedIndex                  = _openedFilesDialog->selectedIndex;
    out.rows.reserve(_openedFilesDialog->rows.size());
    for (const OpenedFileRow& row : _openedFilesDialog->rows)
    {
        OpenedFilesDebugRow debugRow{};
        debugRow.path      = row.path;
        debugRow.file      = row.file;
        debugRow.source    = row.source;
        debugRow.openedBy  = row.openedBy;
        debugRow.pane      = row.pane;
        debugRow.focusable = row.focusable;
        out.rows.push_back(std::move(debugRow));
    }
    return true;
}

bool FolderWindow::DebugSelectOpenedFilesDialogRow(size_t rowIndex) noexcept
{
    if (! _openedFilesDialog || rowIndex >= _openedFilesDialog->rows.size())
    {
        return false;
    }

    return _openedFilesDialog->SelectRow(rowIndex);
}

bool FolderWindow::DebugInvokeOpenedFilesDialogFocusItem() noexcept
{
    return FocusOpenedFilesDialogSelection();
}

void FolderWindow::DebugCloseOpenedFilesDialogForTest() noexcept
{
    CloseOpenedFilesDialog();
}

void FolderWindow::DebugAddOpenedExternalEditorForTest(const std::filesystem::path& path, std::wstring_view openedBy, Pane pane, bool closed) noexcept
{
    if (path.empty())
    {
        return;
    }

    OpenedExternalFileEntry entry{};
    entry.id          = _nextOpenedExternalFileId++;
    entry.source      = OpenedFileSourceKind::Editor;
    entry.path        = path;
    entry.openedBy    = openedBy;
    entry.pane        = pane;
    entry.debugClosed = closed;
    _openedExternalFiles.push_back(std::move(entry));
}

void FolderWindow::DebugClearOpenedExternalEditorsForTest() noexcept
{
    _openedExternalFiles.clear();
}

bool FolderWindow::DebugGetSharedDirectoriesDialogSnapshot(SharedDirectoriesDebugSnapshot& out) const noexcept
{
    out               = {};
    out.selectedIndex = kSharedDirectoriesNoSelection;

    if (! _sharedDirectoriesDialog || ! _sharedDirectoriesDialog->hwnd || IsWindow(_sharedDirectoriesDialog->hwnd.get()) == FALSE)
    {
        return false;
    }

    const HWND dialog                  = _sharedDirectoriesDialog->hwnd.get();
    out.visible                        = IsWindowVisible(dialog) != FALSE;
    out.usesDxUiHost                   = _sharedDirectoriesDialog->dxHost.GetRoot() != nullptr;
    out.visibleChildWindowCount        = CountVisibleChildWindowsLocal(dialog);
    out.visibleNativeChildControlCount = CountVisibleNativeChildControlWindowsLocal(dialog);
    out.dialogClassName                = GetWindowClassNameLocal(dialog);
    out.themeWindowBackground          = _sharedDirectoriesDialog->theme.windowBackground;
    out.themeText                      = _sharedDirectoriesDialog->theme.menu.text;
    out.emptyStateVisible              = _sharedDirectoriesDialog->emptyLabel && _sharedDirectoriesDialog->emptyLabel->IsVisible();
    out.selectedIndex                  = _sharedDirectoriesDialog->selectedIndex;
    out.lastError                      = _sharedDirectoriesDialog->lastError;
    out.rows.reserve(_sharedDirectoriesDialog->rows.size());
    for (const SharedDirectoryRow& row : _sharedDirectoriesDialog->rows)
    {
        SharedDirectoryDebugRow debugRow{};
        debugRow.name      = row.name;
        debugRow.localPath = row.localPath;
        debugRow.type      = row.type;
        debugRow.remark    = row.remark;
        debugRow.openable  = row.openable;
        out.rows.push_back(std::move(debugRow));
    }
    return true;
}

bool FolderWindow::DebugSelectSharedDirectoriesDialogRow(size_t rowIndex) noexcept
{
    if (! _sharedDirectoriesDialog || rowIndex >= _sharedDirectoriesDialog->rows.size())
    {
        return false;
    }

    return _sharedDirectoriesDialog->SelectRow(rowIndex);
}

bool FolderWindow::DebugInvokeSharedDirectoriesDialogOpenPath() noexcept
{
    return OpenSharedDirectoriesDialogSelection();
}

void FolderWindow::DebugCloseSharedDirectoriesDialogForTest() noexcept
{
    CloseSharedDirectoriesDialog();
}

void FolderWindow::DebugSetSharedDirectoriesProviderResultForTest(SharedDirectoriesDebugProviderResult result) noexcept
{
    _debugSharedDirectoriesProviderResult = std::move(result);
    RefreshSharedDirectoriesDialogRows();
}

void FolderWindow::DebugClearSharedDirectoriesProviderForTest() noexcept
{
    _debugSharedDirectoriesProviderResult.reset();
    RefreshSharedDirectoriesDialogRows();
}

bool FolderWindow::DebugHasItemDisplayName(Pane pane, std::wstring_view displayName) const noexcept
{
    const PaneState& state = pane == Pane::Left ? _leftPane : _rightPane;
    return state.folderView.DebugHasItemDisplayName(displayName);
}

size_t FolderWindow::DebugGetItemCount(Pane pane) const noexcept
{
    const PaneState& state = pane == Pane::Left ? _leftPane : _rightPane;
    return state.folderView.DebugGetItemCount();
}

size_t FolderWindow::DebugGetPaneBitmapIconCount(Pane pane) const noexcept
{
    const PaneState& state = pane == Pane::Left ? _leftPane : _rightPane;
    return state.folderView.DebugGetBitmapIconCount();
}

bool FolderWindow::DebugIsItemSelected(Pane pane, std::wstring_view displayName) const noexcept
{
    const PaneState& state = pane == Pane::Left ? _leftPane : _rightPane;
    return state.folderView.DebugIsItemSelectedByDisplayName(displayName);
}

size_t FolderWindow::DebugGetSelectedCount(Pane pane) const noexcept
{
    const PaneState& state = pane == Pane::Left ? _leftPane : _rightPane;
    return state.folderView.DebugGetSelectedItemCount();
}

uint64_t FolderWindow::DebugGetWarmPaneRenderingCallCount(Pane pane) const noexcept
{
    const PaneState& state = pane == Pane::Left ? _leftPane : _rightPane;
    return state.folderView.DebugGetWarmRenderingCallCount();
}

FolderView::DebugWarmPerfSnapshot FolderWindow::DebugGetWarmPanePerfSnapshot(Pane pane) const noexcept
{
    const PaneState& state = pane == Pane::Left ? _leftPane : _rightPane;
    return state.folderView.DebugGetWarmPerfSnapshot();
}

bool FolderWindow::DebugWarmPaneRendering(Pane pane) noexcept
{
    PaneState& state = pane == Pane::Left ? _leftPane : _rightPane;
    return state.folderView.DebugWarmRenderingForSelfTest();
}

bool FolderWindow::DebugGetPaneColumnLayoutSnapshot(Pane pane, FolderView::DebugColumnLayoutSnapshot& out) const
{
    const PaneState& state = pane == Pane::Left ? _leftPane : _rightPane;
    return state.folderView.DebugGetColumnLayoutSnapshot(out);
}

bool FolderWindow::DebugIsEmptyFolderStateActive(Pane pane) const noexcept
{
    const PaneState& state = pane == Pane::Left ? _leftPane : _rightPane;
    return state.folderView.DebugIsEmptyFolderStateActive();
}

std::wstring_view FolderWindow::DebugGetPaneEmptyStateMessage(Pane pane) const noexcept
{
    const PaneState& state = pane == Pane::Left ? _leftPane : _rightPane;
    return state.folderView.DebugGetEmptyStateMessage();
}

std::wstring_view FolderWindow::DebugGetEmptyFolderFunMessage(Pane pane) const noexcept
{
    const PaneState& state = pane == Pane::Left ? _leftPane : _rightPane;
    return state.folderView.DebugGetEmptyFolderFunMessage();
}

FolderView::DebugEmptyFolderItemMetrics FolderWindow::DebugGetEmptyFolderItemMetrics(Pane pane) const noexcept
{
    const PaneState& state = pane == Pane::Left ? _leftPane : _rightPane;
    return state.folderView.DebugGetEmptyFolderItemMetrics();
}

HWND FolderWindow::DebugGetNavigationViewHwnd(Pane pane) const noexcept
{
    const PaneState& state = pane == Pane::Left ? _leftPane : _rightPane;
    return state.hNavigationView.get();
}

bool FolderWindow::DebugGetNavigationViewSnapshot(Pane pane, NavigationViewDebugSnapshot& out) const noexcept
{
    const PaneState& state = pane == Pane::Left ? _leftPane : _rightPane;
    return state.navigationView.DebugGetSnapshot(out);
}

bool FolderWindow::DebugFocusNavigationViewRegion(Pane pane, NavigationView::FocusRegion region) noexcept
{
    PaneState& state = pane == Pane::Left ? _leftPane : _rightPane;
    return state.navigationView.DebugFocusRegion(region);
}

bool FolderWindow::DebugPostCurrentNavigationEditSuggestResult(Pane pane)
{
    PaneState& state = pane == Pane::Left ? _leftPane : _rightPane;
    return state.navigationView.DebugPostCurrentEditSuggestResultForSelfTest();
}

bool FolderWindow::DebugFocusItemByDisplayName(Pane pane, std::wstring_view displayName) noexcept
{
    PaneState& state = pane == Pane::Left ? _leftPane : _rightPane;
    return state.folderView.PrepareForExternalCommand(displayName);
}

bool FolderWindow::DebugGetIncrementalSearchSnapshot(Pane pane, FolderView::IncrementalSearchDebugSnapshot& out) const noexcept
{
    const PaneState& state = pane == Pane::Left ? _leftPane : _rightPane;
    return state.folderView.DebugGetIncrementalSearchSnapshot(out);
}

FolderView::NameFilterState FolderWindow::DebugGetNameFilterState(Pane pane) const
{
    const PaneState& state = pane == Pane::Left ? _leftPane : _rightPane;
    return state.folderView.GetNameFilterState();
}

bool FolderWindow::DebugIsNameFilterActive(Pane pane) const noexcept
{
    const PaneState& state = pane == Pane::Left ? _leftPane : _rightPane;
    return state.folderView.IsNameFilterActive();
}

void FolderWindow::DebugResetPaneVisibilityState(Pane pane) noexcept
{
    PaneState& state = pane == Pane::Left ? _leftPane : _rightPane;
    state.folderView.ShowHiddenNames();
    SetNameFilterState(pane, FolderView::NameFilterState{});
}

FolderView::FilterWatermarkVisualMode FolderWindow::DebugGetFilterWatermarkVisualMode(Pane pane) const noexcept
{
    const PaneState& state = pane == Pane::Left ? _leftPane : _rightPane;
    return state.folderView.DebugGetFilterWatermarkVisualMode();
}
#endif

void FolderWindow::CommandSelectionSelectDialog(Pane pane)
{
    SetActivePane(pane);

    std::vector<std::wstring> history;
    if (_settings && _settings->selectionMasks.has_value())
    {
        history = _settings->selectionMasks->selectHistory;
    }
    MaskSyntax::NormalizeWildcardMaskHistory(history, MaskSyntax::kWildcardMaskHistoryMaxItems);

    HWND ownerWindow = _hWnd ? GetAncestor(_hWnd.get(), GA_ROOT) : nullptr;
    if (! ownerWindow)
    {
        ownerWindow = _hWnd.get();
    }

    const std::optional<std::wstring> maskTextOpt =
        PromptForSelectionMask(ownerWindow, history, _theme, IDS_CAPTION_SELECTION_MASK_SELECT, IDS_LABEL_SELECTION_MASK_SELECT);
    if (! maskTextOpt.has_value())
    {
        return;
    }

    const std::wstring& maskText = maskTextOpt.value();

    if (_settings)
    {
        Common::Settings::SelectionMasksSettings& masks =
            _settings->selectionMasks.has_value() ? _settings->selectionMasks.value() : _settings->selectionMasks.emplace();

        MaskSyntax::AddToWildcardMaskHistory(masks.selectHistory, MaskSyntax::kWildcardMaskHistoryMaxItems, maskText);
        MaskSyntax::NormalizeWildcardMaskHistory(masks.selectHistory, MaskSyntax::kWildcardMaskHistoryMaxItems);
    }

    MaskSyntax::WildcardMask mask = MaskSyntax::ParseWildcardMask(maskText);

    SetPaneSelectionByDisplayNamePredicate(pane, [mask = std::move(mask)](std::wstring_view displayName) noexcept {
        return MaskSyntax::MatchesWildcardMask(displayName, mask);
    }, false /* clearExistingSelection */);
}

void FolderWindow::CommandSelectionUnselectDialog(Pane pane)
{
    SetActivePane(pane);

    std::vector<std::wstring> history;
    if (_settings && _settings->selectionMasks.has_value())
    {
        history = _settings->selectionMasks->unselectHistory;
    }
    MaskSyntax::NormalizeWildcardMaskHistory(history, MaskSyntax::kWildcardMaskHistoryMaxItems);

    HWND ownerWindow = _hWnd ? GetAncestor(_hWnd.get(), GA_ROOT) : nullptr;
    if (! ownerWindow)
    {
        ownerWindow = _hWnd.get();
    }

    const std::optional<std::wstring> maskTextOpt =
        PromptForSelectionMask(ownerWindow, history, _theme, IDS_CAPTION_SELECTION_MASK_UNSELECT, IDS_LABEL_SELECTION_MASK_UNSELECT);
    if (! maskTextOpt.has_value())
    {
        return;
    }

    const std::wstring& maskText = maskTextOpt.value();

    if (_settings)
    {
        Common::Settings::SelectionMasksSettings& masks =
            _settings->selectionMasks.has_value() ? _settings->selectionMasks.value() : _settings->selectionMasks.emplace();

        MaskSyntax::AddToWildcardMaskHistory(masks.unselectHistory, MaskSyntax::kWildcardMaskHistoryMaxItems, maskText);
        MaskSyntax::NormalizeWildcardMaskHistory(masks.unselectHistory, MaskSyntax::kWildcardMaskHistoryMaxItems);
    }

    MaskSyntax::WildcardMask mask = MaskSyntax::ParseWildcardMask(maskText);

    ClearPaneSelectionByDisplayNamePredicate(
        pane, [mask = std::move(mask)](std::wstring_view displayName) noexcept { return MaskSyntax::MatchesWildcardMask(displayName, mask); });
}

void FolderWindow::CommandSelectionInvert(Pane pane)
{
    SetActivePane(pane);
    PaneState& state = pane == Pane::Left ? _leftPane : _rightPane;
    state.folderView.InvertSelection();
}

bool FolderWindow::HasSavedSelection() const noexcept
{
    return _savedSelection.has_value() && ! _savedSelection->displayNames.empty();
}

void FolderWindow::CommandSelectionSave(Pane pane)
{
    SetActivePane(pane);
    PaneState& state = pane == Pane::Left ? _leftPane : _rightPane;

    const std::optional<std::filesystem::path> folderOpt = state.folderView.GetFolderPath();
    if (! folderOpt.has_value() || folderOpt.value().empty())
    {
        MessageBeep(MB_ICONWARNING);
        return;
    }

    std::vector<std::wstring> names = state.folderView.GetSelectedOrFocusedDisplayNames();
    if (names.empty())
    {
        MessageBeep(MB_ICONWARNING);
        return;
    }

    SavedSelection saved{};
    saved.sourcePluginId        = std::wstring(state.folderView.GetFileSystemPluginId());
    saved.sourceInstanceContext = std::wstring(state.folderView.GetFileSystemInstanceContext());
    saved.sourceFolder          = folderOpt.value();
    saved.displayNames          = std::move(names);
    _savedSelection             = std::move(saved);

    std::wstring clipboardText;
    {
        const std::wstring folderText = folderOpt.value().native();
        size_t reserveChars           = folderText.size() + 2u;
        for (const auto& name : _savedSelection->displayNames)
        {
            reserveChars += name.size() + 2u;
        }

        clipboardText.reserve(reserveChars);
        clipboardText.append(folderText);
        for (const auto& name : _savedSelection->displayNames)
        {
            clipboardText.append(L"\r\n");
            clipboardText.append(name);
        }
        clipboardText.append(L"\r\n");
    }

    HWND ownerWindow = _hWnd ? GetAncestor(_hWnd.get(), GA_ROOT) : nullptr;
    if (! ownerWindow)
    {
        ownerWindow = _hWnd.get();
    }

    if (! Common::Clipboard::TrySetUnicodeText(ownerWindow, clipboardText))
    {
        std::wstring title   = LoadStringResource(nullptr, IDS_CAPTION_WARNING);
        std::wstring message = LoadStringResource(nullptr, IDS_MSG_SELECTION_SAVE_CLIPBOARD_FAILED);
        ShowPaneAlertOverlay(
            pane, FolderView::ErrorOverlayKind::Operation, FolderView::OverlaySeverity::Warning, std::move(title), std::move(message), S_OK, true, false);
    }
}

void FolderWindow::CopySelectionText(Pane pane, CopySelectionTextMode mode, UINT titleStringId)
{
    SetActivePane(pane);
    PaneState& state = pane == Pane::Left ? _leftPane : _rightPane;

    const std::vector<std::filesystem::path> paths = state.folderView.GetSelectedOrFocusedPaths();
    if (paths.empty())
    {
        MessageBeep(MB_ICONWARNING);
        return;
    }

    const bool preferUncPath       = mode == CopySelectionTextMode::UncPathAndName && IsFilePluginShortId(state.pluginShortId);
    const auto renderClipboardLine = [mode, preferUncPath](const std::filesystem::path& path) -> std::wstring
    {
        const std::wstring nativePath = path.native();
        switch (mode)
        {
            case CopySelectionTextMode::PathAndName: return nativePath;

            case CopySelectionTextMode::Name:
            {
                std::wstring fileName = path.filename().native();
                return fileName.empty() ? nativePath : fileName;
            }

            case CopySelectionTextMode::Path:
            {
                std::filesystem::path parentPath = path.parent_path();
                if (parentPath.empty() && path.has_root_path())
                {
                    parentPath = path.root_path();
                }

                std::wstring containingPath = parentPath.native();
                return containingPath.empty() ? nativePath : containingPath;
            }

            case CopySelectionTextMode::UncPathAndName: return preferUncPath ? GetUniversalPathOrOriginal(nativePath) : nativePath;
        }

        return nativePath;
    };

    std::vector<std::wstring> lines;
    lines.reserve(paths.size());

    size_t reserveChars = 2u;
    for (const auto& path : paths)
    {
        std::wstring line = renderClipboardLine(path);
        if (line.empty())
        {
            continue;
        }

        reserveChars += line.size() + 2u;
        lines.push_back(std::move(line));
    }

    if (lines.empty())
    {
        MessageBeep(MB_ICONWARNING);
        return;
    }

    std::wstring clipboardText;
    clipboardText.reserve(reserveChars);
    for (size_t i = 0; i < lines.size(); ++i)
    {
        if (i != 0u)
        {
            clipboardText.append(L"\r\n");
        }
        clipboardText.append(lines[i]);
    }
    clipboardText.append(L"\r\n");

    const HWND ownerWindow = GetClipboardOwnerWindow(_hWnd.get());
    if (! Common::Clipboard::TrySetUnicodeText(ownerWindow, clipboardText))
    {
        std::wstring title   = LoadStringResource(nullptr, IDS_CAPTION_WARNING);
        std::wstring message = LoadStringResource(nullptr, IDS_MSG_SELECTION_SAVE_CLIPBOARD_FAILED);
        ShowPaneAlertOverlay(
            pane, FolderView::ErrorOverlayKind::Operation, FolderView::OverlaySeverity::Warning, std::move(title), std::move(message), S_OK, true, false);
        return;
    }

    std::wstring title = LoadStringResource(nullptr, titleStringId);
    if (title.empty())
    {
        title = LoadStringResource(nullptr, IDS_OVERLAY_TITLE_INFORMATION);
    }

    const unsigned long long count = static_cast<unsigned long long>(lines.size());
    const std::wstring_view suffix = count == 1ull ? std::wstring_view(L"") : std::wstring_view(L"s");
    std::wstring message           = FormatStringResource(nullptr, IDS_FMT_COPY_PATH_AND_FILE_NAME_COPIED, count, suffix);
    if (message.empty())
    {
        message = std::format(L"Copied {} item{} to clipboard.", count, suffix);
    }

    ShowPaneAlertOverlay(
        pane, FolderView::ErrorOverlayKind::Operation, FolderView::OverlaySeverity::Information, std::move(title), std::move(message), S_OK, false, false);
}

void FolderWindow::CommandCopyPathAndNameAsText(Pane pane)
{
    CopySelectionText(pane, CopySelectionTextMode::PathAndName, IDS_CMD_COPY_PATH_AND_NAME_AS_TEXT);
}

void FolderWindow::CommandCopyNameAsText(Pane pane)
{
    CopySelectionText(pane, CopySelectionTextMode::Name, IDS_CMD_COPY_NAME_AS_TEXT);
}

void FolderWindow::CommandCopyPathAsText(Pane pane)
{
    CopySelectionText(pane, CopySelectionTextMode::Path, IDS_CMD_COPY_PATH_AS_TEXT);
}

void FolderWindow::CommandCopyUncPathAndNameAsText(Pane pane)
{
    CopySelectionText(pane, CopySelectionTextMode::UncPathAndName, IDS_CMD_COPY_PATH_AND_FILE_NAME);
}

void FolderWindow::CommandMakeFileList(Pane pane)
{
    const auto totalStartedAt = std::chrono::steady_clock::now();
    SetActivePane(pane);
    PaneState& state = pane == Pane::Left ? _leftPane : _rightPane;

    if (state.makeFileListThread.joinable())
    {
        RequestMakeFileListCancellation(pane);
        ShowMakeFileListOverlay(
            *this, pane, FolderView::OverlaySeverity::Information, LoadStringResource(nullptr, IDS_MSG_MAKE_FILE_LIST_CANCELLATION_REQUESTED), S_FALSE);
        Debug::Perf::Emit(L"makeFileList.command_return_us", L"cancel-requested", Debug::Perf::ElapsedUs(totalStartedAt), 0u, 0u, S_FALSE);
        return;
    }

    const std::optional<std::filesystem::path> currentFolder = state.folderView.GetFolderPath();
    if (! state.fileSystem || ! currentFolder.has_value() || currentFolder->empty() || ! IsFilePluginShortId(state.pluginShortId))
    {
        ShowMakeFileListOverlay(
            *this, pane, FolderView::OverlaySeverity::Warning, LoadStringResource(nullptr, IDS_MSG_MAKE_FILE_LIST_LOCAL_FOLDER_REQUIRED), S_FALSE);
        Debug::Perf::Emit(L"makeFileList.total_us", L"unsupported", Debug::Perf::ElapsedUs(totalStartedAt), 0u, 0u, S_FALSE);
        return;
    }

    HWND ownerWindow = _hWnd ? GetAncestor(_hWnd.get(), GA_ROOT) : nullptr;
    if (! ownerWindow)
    {
        ownerWindow = _hWnd.get();
    }

    MakeFileListSettings options = _settings && _settings->makeFileList.has_value() ? _settings->makeFileList.value() : MakeFileListSettings{};
    std::optional<MakeFileListSettings> requestedOptions;
    bool automated = false;
#ifdef ENABLE_TESTS
    if (g_makeFileListAutomation.has_value())
    {
        requestedOptions = g_makeFileListAutomation.value();
        automated        = true;
    }
    else
#endif
    {
        requestedOptions = PromptForMakeFileListOptions(ownerWindow, _theme, options);
    }

    if (! requestedOptions.has_value())
    {
        Debug::Perf::Emit(L"makeFileList.total_us", L"cancelled", Debug::Perf::ElapsedUs(totalStartedAt), 0u, 0u, S_FALSE);
        return;
    }
    options = requestedOptions.value();

    std::vector<std::filesystem::path> selectedPaths;
    if (options.sourceMode == Common::Settings::MakeFileListSourceMode::Selection)
    {
        selectedPaths = state.folderView.GetSelectedOrFocusedPaths();
        if (selectedPaths.empty())
        {
            ShowMakeFileListOverlay(
                *this, pane, FolderView::OverlaySeverity::Warning, LoadStringResource(nullptr, IDS_MSG_MAKE_FILE_LIST_SELECTION_REQUIRED), S_FALSE);
            Debug::Perf::Emit(L"makeFileList.total_us", L"empty-selection", Debug::Perf::ElapsedUs(totalStartedAt), 0u, 0u, S_FALSE);
            return;
        }
    }

    if (options.outputTarget == Common::Settings::MakeFileListOutputTarget::File && ! automated)
    {
        const std::optional<std::filesystem::path> outputFile = PromptForMakeFileListOutputFile(ownerWindow, currentFolder.value(), options);
        if (! outputFile.has_value())
        {
            Debug::Perf::Emit(L"makeFileList.total_us", L"save-cancelled", Debug::Perf::ElapsedUs(totalStartedAt), 0u, 0u, S_FALSE);
            return;
        }
        options.outputFile = outputFile.value();
    }

    if (options.outputTarget == Common::Settings::MakeFileListOutputTarget::File && options.outputFile.empty())
    {
        ShowMakeFileListOverlay(*this,
                                pane,
                                FolderView::OverlaySeverity::Warning,
                                LoadStringResource(nullptr, IDS_MSG_MAKE_FILE_LIST_OUTPUT_FILE_REQUIRED),
                                HRESULT_FROM_WIN32(ERROR_INVALID_NAME));
        Debug::Perf::Emit(
            L"makeFileList.total_us", L"missing-output-file", Debug::Perf::ElapsedUs(totalStartedAt), 0u, 0u, HRESULT_FROM_WIN32(ERROR_INVALID_NAME));
        return;
    }

    InformationalTaskUpdate task{};
    task.kind                        = InformationalTaskUpdate::Kind::MakeFileList;
    task.title                       = LoadStringResource(nullptr, IDS_CMD_MAKE_FILE_LIST);
    task.makeFileListCurrentPath     = currentFolder.value();
    task.makeFileListCollecting      = true;
    const uint64_t taskId            = CreateOrUpdateInformationalTask(task);
    const HWND ownerHwnd             = _hWnd.get();
    const std::filesystem::path root = currentFolder.value();
    const std::wstring title         = task.title;

#ifdef ENABLE_TESTS
    g_makeFileListWorkerActive.store(true, std::memory_order_release);
#endif
    state.makeFileListThread = std::jthread(
        [ownerHwnd, pane, root, selectedPaths = std::move(selectedPaths), options, title, taskId, totalStartedAt](std::stop_token stopToken) noexcept
    {
#ifdef ENABLE_TESTS
        const auto clearWorkerActive = wil::scope_exit([] { g_makeFileListWorkerActive.store(false, std::memory_order_release); });
        const uint32_t delayMs       = g_makeFileListWorkerDelayMs.exchange(0u, std::memory_order_acq_rel);
        uint32_t waitedMs            = 0u;
        while (waitedMs < delayMs && ! stopToken.stop_requested())
        {
            const uint32_t sliceMs = (std::min)(10u, delayMs - waitedMs);
            Sleep(sliceMs);
            waitedMs += sliceMs;
        }
#endif

        MakeFileListProgressState progress{};
        progress.hwnd        = ownerHwnd;
        progress.title       = title;
        progress.taskId      = taskId;
        progress.collecting  = true;
        progress.currentPath = root;
        progress.PostTaskUpdate(false, S_OK);

        uint64_t collectFailures = 0u;
        std::vector<MakeFileListEntry> entries;
        const auto collectStartedAt = std::chrono::steady_clock::now();
        HRESULT operationHr         = CollectMakeFileListEntries(root, selectedPaths, options, stopToken, progress, collectFailures, entries);
        if (operationHr == HRESULT_FROM_WIN32(ERROR_OPERATION_ABORTED) && stopToken.stop_requested())
        {
            operationHr = HRESULT_FROM_WIN32(ERROR_CANCELLED);
        }
        Debug::Perf::Emit(L"makeFileList.collect_us",
                          MakeFileListFormatDetail(options.format),
                          Debug::Perf::ElapsedUs(collectStartedAt),
                          static_cast<uint64_t>(entries.size()),
                          collectFailures,
                          operationHr);

        progress.collecting   = false;
        progress.rendering    = SUCCEEDED(operationHr);
        progress.totalEntries = static_cast<uint64_t>(entries.size());
        progress.MaybePostTaskUpdate();

        std::string outputUtf8;
        std::wstring clipboardText;
        if (SUCCEEDED(operationHr))
        {
            const auto renderStartedAt = std::chrono::steady_clock::now();
            operationHr                = RenderMakeFileListOutput(entries, options, stopToken, progress, outputUtf8, clipboardText);
            if (operationHr == HRESULT_FROM_WIN32(ERROR_OPERATION_ABORTED) && stopToken.stop_requested())
            {
                operationHr = HRESULT_FROM_WIN32(ERROR_CANCELLED);
            }
            Debug::Perf::Emit(L"makeFileList.generate_us",
                              MakeFileListFormatDetail(options.format),
                              Debug::Perf::ElapsedUs(renderStartedAt),
                              static_cast<uint64_t>(entries.size()),
                              static_cast<uint64_t>(outputUtf8.size()),
                              operationHr);
        }

        progress.rendering = false;
        progress.writing   = SUCCEEDED(operationHr);
        progress.MaybePostTaskUpdate();

        uint64_t outputElapsedUs = 0u;
        std::wstring outputTarget;
        if (SUCCEEDED(operationHr))
        {
            if (stopToken.stop_requested())
            {
                operationHr = HRESULT_FROM_WIN32(ERROR_CANCELLED);
            }
            else if (options.outputTarget == Common::Settings::MakeFileListOutputTarget::File)
            {
                const auto outputStartedAt = std::chrono::steady_clock::now();
                operationHr                = WriteMakeFileListUtf8File(options.outputFile, outputUtf8, stopToken);
                if (operationHr == HRESULT_FROM_WIN32(ERROR_OPERATION_ABORTED) && stopToken.stop_requested())
                {
                    operationHr = HRESULT_FROM_WIN32(ERROR_CANCELLED);
                }
                outputElapsedUs = Debug::Perf::ElapsedUs(outputStartedAt);
                outputTarget    = options.outputFile.wstring();
            }
            else
            {
                outputTarget = LoadStringResource(nullptr, IDS_MSG_MAKE_FILE_LIST_TARGET_CLIPBOARD);
            }
        }
        progress.writing = false;

        if (ownerHwnd && IsWindow(ownerHwnd) != FALSE)
        {
            auto completed             = std::make_unique<MakeFileListCompletedPayload>();
            completed->pane            = pane;
            completed->taskId          = taskId;
            completed->title           = title;
            completed->options         = options;
            completed->currentFolder   = root;
            completed->clipboardText   = std::move(clipboardText);
            completed->outputTarget    = std::move(outputTarget);
            completed->entryCount      = static_cast<uint64_t>(entries.size());
            completed->collectFailures = collectFailures;
            completed->outputBytes     = static_cast<uint64_t>(outputUtf8.size());
            completed->outputElapsedUs = outputElapsedUs;
            completed->totalElapsedUs  = Debug::Perf::ElapsedUs(totalStartedAt);
            completed->hr              = operationHr;
            static_cast<void>(PostMessagePayload(ownerHwnd, WndMsg::kMakeFileListCompleted, 0, std::move(completed)));
        }
    });

    Debug::Perf::Emit(L"makeFileList.command_return_us", MakeFileListFormatDetail(options.format), Debug::Perf::ElapsedUs(totalStartedAt), taskId, 0u, S_OK);
}

void FolderWindow::CommandPack(Pane pane)
{
    const auto startedAt = std::chrono::steady_clock::now();
    SetActivePane(pane);
    PaneState& state = pane == Pane::Left ? _leftPane : _rightPane;

    ArchiveOperationResult result{};
#ifdef ENABLE_TESTS
    std::optional<ArchiveCommandDebugOptions> debugOptions = _debugNextArchiveCommandOptions;
    _debugNextArchiveCommandOptions.reset();
    _debugLastArchiveCommandResult.reset();
    const auto recordDebugResult = [&](const ArchiveOperationResult& operationResult)
    {
        ArchiveCommandDebugResult debugResult{};
        debugResult.operation            = L"pack";
        debugResult.hr                   = operationResult.hr;
        debugResult.archivePath          = operationResult.archivePath;
        debugResult.destinationPath      = operationResult.destinationPath;
        debugResult.entryCount           = operationResult.entryCount;
        debugResult.bytesProcessed       = operationResult.bytesProcessed;
        debugResult.skippedConflictCount = operationResult.skippedConflictCount;
        debugResult.entries              = operationResult.entries;
        _debugLastArchiveCommandResult   = std::move(debugResult);
    };
#endif

    const std::optional<std::filesystem::path> currentFolder = state.folderView.GetFolderPath();
    if (! state.fileSystem || ! currentFolder.has_value() || currentFolder->empty() || ! IsFilePluginShortId(state.pluginShortId))
    {
        result.hr = HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
#ifdef ENABLE_TESTS
        recordDebugResult(result);
#endif
        ShowArchiveOverlay(
            *this, pane, IDS_CMD_PACK, FolderView::OverlaySeverity::Warning, LoadStringResource(nullptr, IDS_MSG_ARCHIVE_LOCAL_FOLDER_REQUIRED), result.hr);
        Debug::Perf::Emit(L"archive.pack_us", L"unsupported", Debug::Perf::ElapsedUs(startedAt), 0u, 0u, result.hr);
        return;
    }

    std::vector<std::filesystem::path> selectedPaths = state.folderView.GetSelectedOrFocusedPaths();
    if (selectedPaths.empty())
    {
        result.hr = S_FALSE;
#ifdef ENABLE_TESTS
        recordDebugResult(result);
#endif
        ShowArchiveOverlay(
            *this, pane, IDS_CMD_PACK, FolderView::OverlaySeverity::Warning, LoadStringResource(nullptr, IDS_MSG_ARCHIVE_SELECTION_REQUIRED), result.hr);
        Debug::Perf::Emit(L"archive.pack_us", L"empty-selection", Debug::Perf::ElapsedUs(startedAt), 0u, 0u, result.hr);
        return;
    }

    std::optional<ArchivePackPromptResult> packPromptResult;
    std::optional<std::filesystem::path> archivePath;
    ArchivePackerDefinition selectedPacker = BuildStoredZipPackerDefinition();
    bool deleteSourcesAfterPack            = false;
    bool overwrite                         = true;
#ifdef ENABLE_TESTS
    if (debugOptions.has_value())
    {
        archivePath    = debugOptions->archivePath;
        overwrite      = debugOptions->overwrite;
        selectedPacker = BuildStoredZipPackerDefinition();
    }
    else
#endif
    {
        HWND ownerWindow = _hWnd ? GetAncestor(_hWnd.get(), GA_ROOT) : nullptr;
        if (! ownerWindow)
        {
            ownerWindow = _hWnd.get();
        }
        packPromptResult = PromptForArchivePackOptions(ownerWindow, _theme, currentFolder.value(), selectedPaths);
        if (packPromptResult.has_value())
        {
            archivePath            = packPromptResult->archivePath;
            selectedPacker         = packPromptResult->packer;
            deleteSourcesAfterPack = packPromptResult->deleteSources;
            overwrite              = false;
        }
    }

    if (! archivePath.has_value())
    {
        result.hr = S_FALSE;
#ifdef ENABLE_TESTS
        recordDebugResult(result);
#endif
        Debug::Perf::Emit(L"archive.pack_us", L"cancelled", Debug::Perf::ElapsedUs(startedAt), 0u, 0u, result.hr);
        return;
    }

    if (archivePath->empty())
    {
        result.archivePath = archivePath.value();
        result.hr          = HRESULT_FROM_WIN32(ERROR_INVALID_NAME);
#ifdef ENABLE_TESTS
        recordDebugResult(result);
#endif
        ShowArchiveOverlay(
            *this, pane, IDS_CMD_PACK, FolderView::OverlaySeverity::Warning, LoadStringResource(nullptr, IDS_MSG_ARCHIVE_ARCHIVE_PATH_REQUIRED), result.hr);
        Debug::Perf::Emit(L"archive.pack_us", L"missing-archive", Debug::Perf::ElapsedUs(startedAt), 0u, 0u, result.hr);
        return;
    }

    if (selectedPacker.extensionNoDot.empty())
    {
        selectedPacker = BuildStoredZipPackerDefinition();
    }

    *archivePath = EnsureArchivePathExtension(archivePath.value(), selectedPacker.extensionNoDot);
    if (archivePath->extension().empty())
    {
        archivePath->replace_extension(L".zip");
    }

    if (const HRESULT safetyHr = ValidateArchiveOutputOutsideSelectedSources(selectedPaths, archivePath.value()); FAILED(safetyHr))
    {
        result.archivePath = archivePath.value();
        result.hr          = safetyHr;
#ifdef ENABLE_TESTS
        recordDebugResult(result);
#endif
        ShowArchiveOverlay(
            *this, pane, IDS_CMD_PACK, FolderView::OverlaySeverity::Warning, LoadStringResource(nullptr, IDS_MSG_ARCHIVE_OUTPUT_INSIDE_SOURCE), result.hr);
        Debug::Perf::Emit(L"archive.pack_us", L"unsafe-output", Debug::Perf::ElapsedUs(startedAt), 0u, 0u, result.hr);
        return;
    }

    if (deleteSourcesAfterPack)
    {
        HWND ownerWindow = _hWnd ? GetAncestor(_hWnd.get(), GA_ROOT) : nullptr;
        if (! ownerWindow)
        {
            ownerWindow = _hWnd.get();
        }

        if (! ConfirmPermanentDeletePaths(ownerWindow, selectedPaths))
        {
            result.archivePath = archivePath.value();
            result.hr          = S_FALSE;
#ifdef ENABLE_TESTS
            recordDebugResult(result);
#endif
            Debug::Perf::Emit(L"archive.pack_us", L"delete-cancelled", Debug::Perf::ElapsedUs(startedAt), 0u, 0u, result.hr);
            return;
        }
    }

    result = selectedPacker.storedZip ? CreateStoredZipArchive(selectedPaths, archivePath.value(), overwrite)
                                      : CreateSevenZipArchive(selectedPaths, archivePath.value(), overwrite, selectedPacker);
#ifdef ENABLE_TESTS
    recordDebugResult(result);
#endif
    Debug::Perf::Emit(
        L"archive.pack_us", selectedPacker.extensionNoDot, Debug::Perf::ElapsedUs(startedAt), result.entryCount, result.bytesProcessed, result.hr);

    if (FAILED(result.hr))
    {
        ShowArchiveOverlay(*this,
                           pane,
                           IDS_CMD_PACK,
                           FolderView::OverlaySeverity::Warning,
                           ArchiveFailureMessage(IDS_FMT_ARCHIVE_PACK_FAILED, result.archivePath, result.hr),
                           result.hr);
        return;
    }

    if (deleteSourcesAfterPack)
    {
        FolderView::FileOperationRequest deleteRequest{};
        deleteRequest.operation   = FILESYSTEM_DELETE;
        deleteRequest.sourcePaths = selectedPaths;
        deleteRequest.flags       = static_cast<FileSystemFlags>(FILESYSTEM_FLAG_RECURSIVE);
        const HRESULT deleteHr    = StartFileOperationFromFolderView(pane, std::move(deleteRequest));
        if (FAILED(deleteHr))
        {
            ShowArchiveOverlay(
                *this,
                pane,
                IDS_CMD_PACK,
                FolderView::OverlaySeverity::Warning,
                FormatStringResource(
                    nullptr, IDS_FMT_ARCHIVE_DELETE_SOURCES_FAILED, result.archivePath.wstring(), static_cast<unsigned long>(static_cast<uint32_t>(deleteHr))),
                deleteHr);
            RefreshFolderViewIfPathMatches(state.folderView, result.archivePath.parent_path());
            return;
        }
    }

    RefreshFolderViewIfPathMatches(state.folderView, result.archivePath.parent_path());
    ShowArchiveOverlay(*this,
                       pane,
                       IDS_CMD_PACK,
                       FolderView::OverlaySeverity::Information,
                       FormatStringResource(nullptr, IDS_FMT_ARCHIVE_PACK_COMPLETED, result.archivePath.wstring(), result.entryCount),
                       S_OK);
}

void FolderWindow::CommandUnpack(Pane pane)
{
    const auto startedAt = std::chrono::steady_clock::now();
    SetActivePane(pane);
    PaneState& state = pane == Pane::Left ? _leftPane : _rightPane;

    ArchiveOperationResult result{};
#ifdef ENABLE_TESTS
    std::optional<ArchiveCommandDebugOptions> debugOptions = _debugNextArchiveCommandOptions;
    _debugNextArchiveCommandOptions.reset();
    _debugLastArchiveCommandResult.reset();
    const auto recordDebugResult = [&](const ArchiveOperationResult& operationResult)
    {
        ArchiveCommandDebugResult debugResult{};
        debugResult.operation            = L"unpack";
        debugResult.hr                   = operationResult.hr;
        debugResult.archivePath          = operationResult.archivePath;
        debugResult.destinationPath      = operationResult.destinationPath;
        debugResult.entryCount           = operationResult.entryCount;
        debugResult.bytesProcessed       = operationResult.bytesProcessed;
        debugResult.skippedConflictCount = operationResult.skippedConflictCount;
        debugResult.entries              = operationResult.entries;
        _debugLastArchiveCommandResult   = std::move(debugResult);
    };
#endif

    const std::optional<std::filesystem::path> currentFolder = state.folderView.GetFolderPath();
    if (! state.fileSystem || ! currentFolder.has_value() || currentFolder->empty() || ! IsFilePluginShortId(state.pluginShortId))
    {
        result.hr = HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
#ifdef ENABLE_TESTS
        recordDebugResult(result);
#endif
        ShowArchiveOverlay(
            *this, pane, IDS_CMD_UNPACK, FolderView::OverlaySeverity::Warning, LoadStringResource(nullptr, IDS_MSG_ARCHIVE_LOCAL_FOLDER_REQUIRED), result.hr);
        Debug::Perf::Emit(L"archive.unpack_us", L"unsupported", Debug::Perf::ElapsedUs(startedAt), 0u, 0u, result.hr);
        return;
    }

    std::vector<std::filesystem::path> selectedPaths = state.folderView.GetSelectedOrFocusedPaths();
    if (selectedPaths.empty())
    {
        result.hr = S_FALSE;
#ifdef ENABLE_TESTS
        recordDebugResult(result);
#endif
        ShowArchiveOverlay(
            *this, pane, IDS_CMD_UNPACK, FolderView::OverlaySeverity::Warning, LoadStringResource(nullptr, IDS_MSG_ARCHIVE_SELECTION_REQUIRED), result.hr);
        Debug::Perf::Emit(L"archive.unpack_us", L"empty-selection", Debug::Perf::ElapsedUs(startedAt), 0u, 0u, result.hr);
        return;
    }

    std::optional<std::filesystem::path> destinationPath;
    ArchiveUnpackerDefinition selectedUnpacker = BuildStoredZipUnpackerDefinition();
    bool deleteArchiveAfterUnpack              = false;
    std::wstring maskText                      = L"*.*";
    ArchiveExistingTargetPolicy conflictPolicy = ArchiveExistingTargetPolicy::Skip;
#ifdef ENABLE_TESTS
    if (debugOptions.has_value())
    {
        destinationPath  = debugOptions->destinationPath;
        conflictPolicy   = debugOptions->overwrite ? ArchiveExistingTargetPolicy::Replace : ArchiveExistingTargetPolicy::Skip;
        selectedUnpacker = BuildStoredZipUnpackerDefinition();
    }
    else
#endif
    {
        HWND ownerWindow = _hWnd ? GetAncestor(_hWnd.get(), GA_ROOT) : nullptr;
        if (! ownerWindow)
        {
            ownerWindow = _hWnd.get();
        }
        const std::optional<ArchiveUnpackPromptResult> unpackPromptResult =
            PromptForArchiveUnpackOptions(ownerWindow, _theme, currentFolder.value(), selectedPaths.front());
        if (unpackPromptResult.has_value())
        {
            destinationPath          = unpackPromptResult->destinationPath;
            selectedUnpacker         = unpackPromptResult->unpacker;
            deleteArchiveAfterUnpack = unpackPromptResult->deleteArchive;
            maskText                 = unpackPromptResult->maskText;
            conflictPolicy           = unpackPromptResult->conflictPolicy;
        }
    }

    if (! destinationPath.has_value())
    {
        result.hr = S_FALSE;
#ifdef ENABLE_TESTS
        recordDebugResult(result);
#endif
        Debug::Perf::Emit(L"archive.unpack_us", L"cancelled", Debug::Perf::ElapsedUs(startedAt), 0u, 0u, result.hr);
        return;
    }

    if (destinationPath->empty())
    {
        result.archivePath     = selectedPaths.front();
        result.destinationPath = destinationPath.value();
        result.hr              = HRESULT_FROM_WIN32(ERROR_INVALID_NAME);
#ifdef ENABLE_TESTS
        recordDebugResult(result);
#endif
        ShowArchiveOverlay(
            *this, pane, IDS_CMD_UNPACK, FolderView::OverlaySeverity::Warning, LoadStringResource(nullptr, IDS_MSG_ARCHIVE_DESTINATION_REQUIRED), result.hr);
        Debug::Perf::Emit(L"archive.unpack_us", L"missing-destination", Debug::Perf::ElapsedUs(startedAt), 0u, 0u, result.hr);
        return;
    }

    if (! selectedUnpacker.storedZip)
    {
        result.archivePath     = selectedPaths.front();
        result.destinationPath = destinationPath.value();
        result.hr              = HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
#ifdef ENABLE_TESTS
        recordDebugResult(result);
#endif
        ShowArchiveOverlay(
            *this, pane, IDS_CMD_UNPACK, FolderView::OverlaySeverity::Warning, LoadStringResource(nullptr, IDS_MSG_ARCHIVE_UNSUPPORTED_FORMAT), result.hr);
        Debug::Perf::Emit(L"archive.unpack_us", L"unsupported-unpacker", Debug::Perf::ElapsedUs(startedAt), 0u, 0u, result.hr);
        return;
    }

    if (deleteArchiveAfterUnpack)
    {
        HWND ownerWindow = _hWnd ? GetAncestor(_hWnd.get(), GA_ROOT) : nullptr;
        if (! ownerWindow)
        {
            ownerWindow = _hWnd.get();
        }

        if (! ConfirmPermanentDeletePaths(ownerWindow, selectedPaths))
        {
            result.archivePath     = selectedPaths.front();
            result.destinationPath = destinationPath.value();
            result.hr              = S_FALSE;
#ifdef ENABLE_TESTS
            recordDebugResult(result);
#endif
            Debug::Perf::Emit(L"archive.unpack_us", L"delete-cancelled", Debug::Perf::ElapsedUs(startedAt), 0u, 0u, result.hr);
            return;
        }
    }

    result.destinationPath = destinationPath.value();
    for (const std::filesystem::path& archivePath : selectedPaths)
    {
        ArchiveOperationResult currentResult = ExtractStoredZipArchive(archivePath, destinationPath.value(), conflictPolicy, maskText);
        if (currentResult.hr == HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED) && PathMatchSpecW(archivePath.c_str(), L"*.zip"))
        {
            currentResult = ExtractSevenZipArchive(archivePath, destinationPath.value(), conflictPolicy, maskText);
        }
        if (FAILED(currentResult.hr))
        {
            const std::filesystem::path failedArchivePath = currentResult.archivePath;
            result.hr                                     = currentResult.hr;
            result.destinationPath                        = destinationPath.value();
            result.entryCount += currentResult.entryCount;
            result.bytesProcessed += currentResult.bytesProcessed;
            result.skippedConflictCount += currentResult.skippedConflictCount;
            result.entries.insert(result.entries.end(), currentResult.entries.begin(), currentResult.entries.end());
#ifdef ENABLE_TESTS
            recordDebugResult(result);
#endif
            Debug::Perf::Emit(
                L"archive.unpack_us", selectedUnpacker.extensionNoDot, Debug::Perf::ElapsedUs(startedAt), result.entryCount, result.bytesProcessed, result.hr);
            const UINT messageId = result.hr == HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED) ? IDS_MSG_ARCHIVE_UNSUPPORTED_FORMAT : IDS_FMT_ARCHIVE_UNPACK_FAILED;
            const std::wstring message = messageId == IDS_MSG_ARCHIVE_UNSUPPORTED_FORMAT
                                             ? LoadStringResource(nullptr, IDS_MSG_ARCHIVE_UNSUPPORTED_FORMAT)
                                             : ArchiveFailureMessage(IDS_FMT_ARCHIVE_UNPACK_FAILED, failedArchivePath, result.hr);
            ShowArchiveOverlay(*this, pane, IDS_CMD_UNPACK, FolderView::OverlaySeverity::Warning, message, result.hr);
            return;
        }

        if (result.archivePath.empty())
        {
            result.archivePath = currentResult.archivePath;
        }
        result.entryCount += currentResult.entryCount;
        result.bytesProcessed += currentResult.bytesProcessed;
        result.skippedConflictCount += currentResult.skippedConflictCount;
        result.entries.insert(result.entries.end(), currentResult.entries.begin(), currentResult.entries.end());
    }

    result.hr = S_OK;
#ifdef ENABLE_TESTS
    recordDebugResult(result);
#endif
    Debug::Perf::Emit(
        L"archive.unpack_us", selectedUnpacker.extensionNoDot, Debug::Perf::ElapsedUs(startedAt), result.entryCount, result.bytesProcessed, result.hr);

    if (deleteArchiveAfterUnpack)
    {
        const HRESULT deleteHr = DeleteUnpackedArchives(selectedPaths);
        if (FAILED(deleteHr))
        {
            ShowArchiveOverlay(
                *this,
                pane,
                IDS_CMD_UNPACK,
                FolderView::OverlaySeverity::Warning,
                FormatStringResource(
                    nullptr, IDS_FMT_ARCHIVE_DELETE_ARCHIVE_FAILED, result.archivePath.wstring(), static_cast<unsigned long>(static_cast<uint32_t>(deleteHr))),
                deleteHr);
            RefreshFolderViewIfPathMatches(state.folderView, result.destinationPath);
            return;
        }
    }

    RefreshFolderViewIfPathMatches(state.folderView, result.destinationPath);
    for (const std::filesystem::path& archivePath : selectedPaths)
    {
        RefreshFolderViewIfPathMatches(state.folderView, archivePath.parent_path());
    }
    ShowArchiveOverlay(*this,
                       pane,
                       IDS_CMD_UNPACK,
                       FolderView::OverlaySeverity::Information,
                       selectedPaths.size() == 1u
                           ? FormatStringResource(nullptr, IDS_FMT_ARCHIVE_UNPACK_COMPLETED, result.entryCount, result.archivePath.wstring())
                           : FormatStringResource(nullptr, IDS_FMT_ARCHIVE_UNPACK_COMPLETED_MULTIPLE, result.entryCount, selectedPaths.size()),
                       S_OK);
}

#ifdef ENABLE_TESTS
void FolderWindow::DebugSetNextArchiveCommandOptionsForTest(ArchiveCommandDebugOptions options) noexcept
{
    _debugNextArchiveCommandOptions = std::move(options);
}

void FolderWindow::DebugClearArchiveCommandOptionsForTest() noexcept
{
    _debugNextArchiveCommandOptions.reset();
    _debugLastArchiveCommandResult.reset();
}

std::optional<FolderWindow::ArchiveCommandDebugResult> FolderWindow::DebugGetLastArchiveCommandResultForTest() const noexcept
{
    return _debugLastArchiveCommandResult;
}
#endif

void FolderWindow::RegisterOpenedExternalFile(OpenedFileSourceKind source,
                                              const std::filesystem::path& path,
                                              std::wstring_view openedBy,
                                              Pane pane,
                                              wil::unique_handle processHandle,
                                              DWORD processId) noexcept
{
    if (path.empty())
    {
        return;
    }

    OpenedExternalFileEntry entry{};
    entry.id            = _nextOpenedExternalFileId++;
    entry.source        = source;
    entry.path          = path;
    entry.openedBy      = openedBy;
    entry.pane          = pane;
    entry.processId     = processId;
    entry.processHandle = std::move(processHandle);
    _openedExternalFiles.push_back(std::move(entry));
}

void FolderWindow::PruneClosedOpenedExternalFiles() noexcept
{
    for (auto it = _openedExternalFiles.begin(); it != _openedExternalFiles.end();)
    {
        bool remove = false;
#ifdef ENABLE_TESTS
        if (it->debugClosed)
        {
            remove = true;
        }
#endif
        if (! remove && it->processHandle)
        {
            const DWORD waitResult = WaitForSingleObject(it->processHandle.get(), 0u);
            remove                 = waitResult == WAIT_OBJECT_0 || waitResult == WAIT_FAILED;
        }

        if (remove)
        {
            it = _openedExternalFiles.erase(it);
        }
        else
        {
            ++it;
        }
    }
}

std::wstring FolderWindow::ResolveOpenedFilesViewerName(const ViewerInstance& instance) const
{
    if (! instance.openedBy.empty())
    {
        return instance.openedBy;
    }

    const auto& plugins = ViewerPluginManager::GetInstance().GetPlugins();
    for (const ViewerPluginManager::PluginEntry& plugin : plugins)
    {
        if (EqualsNoCase(plugin.id, instance.viewerPluginId))
        {
            if (! plugin.name.empty())
            {
                return plugin.name;
            }
            if (! plugin.shortId.empty())
            {
                return plugin.shortId;
            }
            break;
        }
    }

    std::wstring fallback = LoadStringResource(nullptr, IDS_OPENED_FILES_UNKNOWN_VIEWER);
    if (fallback.empty())
    {
        fallback = instance.viewerPluginId.empty() ? std::wstring(L"Viewer") : instance.viewerPluginId;
    }
    return fallback;
}

std::vector<FolderWindow::OpenedFileRow> FolderWindow::CollectOpenedFileRows() noexcept
{
    Debug::Perf::Scope perf(L"listOpenedFiles.collect_us");
    PruneClosedOpenedExternalFiles();

    std::vector<OpenedFileRow> rows;
    rows.reserve(_openedExternalFiles.size() + _viewerInstances.size() + (_previewSourcePane.has_value() ? 1u : 0u));

    const auto sourceLabel = [](OpenedFileSourceKind source) -> std::wstring
    {
        switch (source)
        {
            case OpenedFileSourceKind::Viewer: return LoadStringOrFallback(IDS_OPENED_FILES_SOURCE_VIEWER, L"Viewer");
            case OpenedFileSourceKind::Editor: return LoadStringOrFallback(IDS_OPENED_FILES_SOURCE_EDITOR, L"Editor");
            case OpenedFileSourceKind::Preview: return LoadStringOrFallback(IDS_OPENED_FILES_SOURCE_PREVIEW, L"Preview Pane");
        }
        return LoadStringOrFallback(IDS_OPENED_FILES_SOURCE_VIEWER, L"Viewer");
    };

    for (const OpenedExternalFileEntry& entry : _openedExternalFiles)
    {
        if (entry.path.empty())
        {
            continue;
        }

        OpenedFileRow row{};
        row.path      = entry.path;
        row.file      = OpenedFilesDisplayNameForPath(entry.path);
        row.source    = sourceLabel(entry.source);
        row.openedBy  = entry.openedBy.empty() ? row.source : entry.openedBy;
        row.pane      = entry.pane;
        row.focusable = true;
        rows.push_back(std::move(row));
    }

    for (const auto& instance : _viewerInstances)
    {
        if (! instance || instance->focusedPath.empty() || instance->source == OpenedFileSourceKind::Preview)
        {
            continue;
        }

        const std::filesystem::path path(instance->focusedPath);
        if (path.empty())
        {
            continue;
        }

        OpenedFileRow row{};
        row.path      = path;
        row.file      = OpenedFilesDisplayNameForPath(path);
        row.source    = sourceLabel(OpenedFileSourceKind::Viewer);
        row.openedBy  = ResolveOpenedFilesViewerName(*instance);
        row.pane      = instance->pane;
        row.focusable = true;
        rows.push_back(std::move(row));
    }

    if (_previewSourcePane.has_value())
    {
        const Pane sourcePane      = _previewSourcePane.value();
        const Pane hostPane        = OppositePane(sourcePane);
        const PaneState& hostState = hostPane == Pane::Left ? _leftPane : _rightPane;
        if (! hostState.previewedPath.empty())
        {
            OpenedFileRow row{};
            row.path      = hostState.previewedPath;
            row.file      = OpenedFilesDisplayNameForPath(hostState.previewedPath);
            row.source    = sourceLabel(OpenedFileSourceKind::Preview);
            row.openedBy  = LoadStringOrFallback(IDS_OPENED_FILES_OPENED_BY_PREVIEW, L"Preview Pane");
            row.pane      = sourcePane;
            row.focusable = true;
            rows.push_back(std::move(row));
        }
    }

    perf.SetValue0(static_cast<uint64_t>(rows.size()));
    perf.SetValue1(static_cast<uint64_t>(_openedExternalFiles.size()));
    return rows;
}

void FolderWindow::RefreshOpenedFilesDialogRows() noexcept
{
    if (! _openedFilesDialog || ! _openedFilesDialog->hwnd || IsWindow(_openedFilesDialog->hwnd.get()) == FALSE)
    {
        return;
    }

    _openedFilesDialog->SetRows(CollectOpenedFileRows());
}

void FolderWindow::ApplyOpenedFilesDialogTheme() noexcept
{
    if (! _openedFilesDialog || ! _openedFilesDialog->hwnd || IsWindow(_openedFilesDialog->hwnd.get()) == FALSE)
    {
        return;
    }

    _openedFilesDialog->theme = _theme;
    _openedFilesDialog->ApplyTheme();
}

void FolderWindow::RequestCloseOpenedFilesDialog() noexcept
{
    if (_hWnd && IsWindow(_hWnd.get()) != FALSE && PostMessageW(_hWnd.get(), WndMsg::kFolderWindowCloseOpenedFilesDialog, 0, 0) != 0)
    {
        return;
    }

    CloseOpenedFilesDialog();
}

void FolderWindow::CloseOpenedFilesDialog() noexcept
{
    if (! _openedFilesDialog)
    {
        return;
    }

    if (_openedFilesDialog->hwnd && ! _openedFilesDialog->destroyed && IsWindow(_openedFilesDialog->hwnd.get()) != FALSE)
    {
        _openedFilesDialog->hwnd.reset();
    }
    _openedFilesDialog.reset();
}

bool FolderWindow::FocusOpenedFileRow(const OpenedFileRow& row) noexcept
{
    if (! row.focusable || row.path.empty())
    {
        return false;
    }

    std::filesystem::path folderPath = row.path.parent_path();
    if (folderPath.empty() && row.path.has_root_path())
    {
        folderPath = row.path.root_path();
    }
    if (folderPath.empty())
    {
        return false;
    }

    std::wstring focusName = row.path.filename().wstring();
    if (focusName.empty())
    {
        focusName = row.file;
    }

    SetActivePane(row.pane);
    return SUCCEEDED(ExecuteInActivePane(folderPath, focusName, 0u, true));
}

bool FolderWindow::FocusOpenedFilesDialogSelection() noexcept
{
    if (! _openedFilesDialog)
    {
        return false;
    }

    if (_openedFilesDialog->selectedIndex == kOpenedFilesNoSelection || _openedFilesDialog->selectedIndex >= _openedFilesDialog->rows.size())
    {
        return false;
    }

    const OpenedFileRow row = _openedFilesDialog->rows[_openedFilesDialog->selectedIndex];
    const bool focused      = FocusOpenedFileRow(row);
    if (focused)
    {
        RequestCloseOpenedFilesDialog();
    }
    return focused;
}

void FolderWindow::CommandListOpenedFiles(Pane pane)
{
    Debug::Perf::Scope perf(L"listOpenedFiles.open_us");
    SetActivePane(pane);

    if (_openedFilesDialog && _openedFilesDialog->hwnd && IsWindow(_openedFilesDialog->hwnd.get()) != FALSE)
    {
        _openedFilesDialog->pane  = pane;
        _openedFilesDialog->theme = _theme;
        _openedFilesDialog->ApplyTheme();
        RefreshOpenedFilesDialogRows();
        ShowWindow(_openedFilesDialog->hwnd.get(), SW_SHOWNORMAL);
        SetForegroundWindow(_openedFilesDialog->hwnd.get());
        perf.SetValue0(static_cast<uint64_t>(_openedFilesDialog->rows.size()));
        return;
    }

    std::unique_ptr<OpenedFilesDialogState, OpenedFilesDialogStateDeleter> state{std::make_unique<OpenedFilesDialogState>().release()};
    state->owner         = this;
    state->pane          = pane;
    state->theme         = _theme;
    state->rows          = CollectOpenedFileRows();
    state->selectedIndex = state->rows.empty() ? kOpenedFilesNoSelection : 0u;

    HWND ownerWindow = _hWnd ? GetAncestor(_hWnd.get(), GA_ROOT) : nullptr;
    if (! ownerWindow)
    {
        ownerWindow = _hWnd.get();
    }

    const HRESULT classHr = OpenedFilesDialogState::EnsureWindowClass();
    if (FAILED(classHr))
    {
        perf.SetHr(classHr);
        Debug::Warning(L"FolderWindow::CommandListOpenedFiles: failed to register opened files window class (hr=0x{:08X}).",
                       static_cast<unsigned long>(classHr));
        return;
    }

    const DWORD style   = WS_CAPTION | WS_SYSMENU | WS_POPUP | WS_THICKFRAME | WS_CLIPCHILDREN | WS_CLIPSIBLINGS;
    const DWORD exStyle = WS_EX_DLGMODALFRAME;
    const UINT dpi      = ownerWindow && IsWindow(ownerWindow) != FALSE ? GetDpiForWindow(ownerWindow) : GetDpiForSystem();
    RECT bounds{0, 0, ScalePanePromptForDpi(dpi, 720), ScalePanePromptForDpi(dpi, 460)};
    if (AdjustWindowRectExForDpi(&bounds, style, FALSE, exStyle, dpi) == FALSE)
    {
        const HRESULT hr = HRESULT_FROM_WIN32(GetLastError());
        perf.SetHr(hr);
        Debug::Warning(L"FolderWindow::CommandListOpenedFiles: failed to calculate opened files window bounds (hr=0x{:08X}).", static_cast<unsigned long>(hr));
        return;
    }

    const std::wstring caption = LoadStringResource(nullptr, IDS_CMD_LIST_OPENED_FILES);
    HWND dialog                = CreateWindowExW(exStyle,
                                                 kOpenedFilesWindowClassName,
                                                 caption.c_str(),
                                                 style,
                                                 CW_USEDEFAULT,
                                                 CW_USEDEFAULT,
                                                 bounds.right - bounds.left,
                                                 bounds.bottom - bounds.top,
                                                 ownerWindow,
                                                 nullptr,
                                                 GetModuleHandleW(nullptr),
                                                 state.get());
    if (! dialog)
    {
        perf.SetHr(HRESULT_FROM_WIN32(GetLastError()));
        Debug::Warning(L"FolderWindow::CommandListOpenedFiles: failed to create opened files dialog.");
        return;
    }

    if (! state->hwnd)
    {
        state->hwnd = wil::unique_hwnd(dialog);
    }
    perf.SetValue0(static_cast<uint64_t>(state->rows.size()));
    _openedFilesDialog = std::move(state);
    static_cast<void>(Common::WindowSizing::CenterExistingWindowOnOwner(dialog, ownerWindow));
    static_cast<void>(_openedFilesDialog->dxHost.PrimeForShow());
    ShowWindow(dialog, SW_SHOWNORMAL);
    UpdateWindow(dialog);
    SetForegroundWindow(dialog);
}

std::vector<FolderWindow::SharedDirectoryRow> FolderWindow::CollectSharedDirectoryRows(HRESULT& outHr) const noexcept
{
    Debug::Perf::Scope perf(L"sharedDirectories.collect_us");
    outHr = S_OK;

    const auto sortRows = [](std::vector<SharedDirectoryRow>& rows) noexcept
    {
        std::sort(rows.begin(),
                  rows.end(),
                  [](const SharedDirectoryRow& left, const SharedDirectoryRow& right) noexcept
        {
            const int nameResult =
                CompareStringOrdinal(left.name.data(), static_cast<int>(left.name.size()), right.name.data(), static_cast<int>(right.name.size()), TRUE);
            if (nameResult != CSTR_EQUAL)
            {
                return nameResult == CSTR_LESS_THAN;
            }

            const int pathResult = CompareStringOrdinal(
                left.localPath.data(), static_cast<int>(left.localPath.size()), right.localPath.data(), static_cast<int>(right.localPath.size()), TRUE);
            return pathResult == CSTR_LESS_THAN;
        });
    };

#ifdef ENABLE_TESTS
    if (_debugSharedDirectoriesProviderResult.has_value())
    {
        const SharedDirectoriesDebugProviderResult& providerResult = _debugSharedDirectoriesProviderResult.value();
        outHr                                                      = providerResult.hr;
        if (FAILED(outHr))
        {
            perf.SetHr(outHr);
            return {};
        }

        std::vector<SharedDirectoryRow> rows;
        rows.reserve(providerResult.rows.size());
        for (const SharedDirectoryDebugRow& source : providerResult.rows)
        {
            SharedDirectoryRow row{};
            row.name      = source.name;
            row.localPath = source.localPath;
            row.type      = source.type.empty() ? LoadStringOrFallback(IDS_SHARED_DIRECTORIES_TYPE_DISK, L"Disk") : source.type;
            row.remark    = source.remark;
            row.openable  = source.openable;
            rows.push_back(std::move(row));
        }
        sortRows(rows);
        perf.SetValue0(static_cast<uint64_t>(rows.size()));
        perf.SetHr(outHr);
        return rows;
    }
#endif

    std::vector<SharedDirectoryRow> rows;
    DWORD resumeHandle = 0;

    for (;;)
    {
        LPBYTE shareBufferRaw       = nullptr;
        DWORD entriesRead           = 0;
        DWORD totalEntries          = 0;
        const NET_API_STATUS status = ::NetShareEnum(nullptr, 2, &shareBufferRaw, MAX_PREFERRED_LENGTH, &entriesRead, &totalEntries, &resumeHandle);

        wil::unique_any<LPBYTE, decltype(&::NetApiBufferFree), ::NetApiBufferFree> shareBuffer(shareBufferRaw);

        if (status != NERR_Success && status != ERROR_MORE_DATA)
        {
            outHr = HRESULT_FROM_WIN32(status);
            perf.SetHr(outHr);
            Debug::Warning(L"Shared Directories: NetShareEnum failed for local machine (hr=0x{:08X}).", static_cast<unsigned long>(outHr));
            return {};
        }

        const auto* shareInfo = reinterpret_cast<const SHARE_INFO_2*>(shareBuffer.get());
        for (DWORD index = 0; index < entriesRead; ++index)
        {
            const SHARE_INFO_2& entry = shareInfo[index];
            if (! entry.shi2_netname || entry.shi2_netname[0] == L'\0')
            {
                continue;
            }

            const DWORD shareType = entry.shi2_type & STYPE_MASK;
            if (shareType != STYPE_DISKTREE)
            {
                continue;
            }

            SharedDirectoryRow row{};
            row.name      = entry.shi2_netname;
            row.localPath = entry.shi2_path ? entry.shi2_path : L"";
            row.type      = LoadStringOrFallback(IDS_SHARED_DIRECTORIES_TYPE_DISK, L"Disk");
            row.remark    = entry.shi2_remark ? entry.shi2_remark : L"";
            row.openable  = SharedDirectoryPathExists(row.localPath);
            rows.push_back(std::move(row));
        }

        if (status == NERR_Success)
        {
            break;
        }
    }

    sortRows(rows);
    perf.SetValue0(static_cast<uint64_t>(rows.size()));
    perf.SetHr(outHr);
    return rows;
}

void FolderWindow::RefreshSharedDirectoriesDialogRows() noexcept
{
    if (! _sharedDirectoriesDialog || ! _sharedDirectoriesDialog->hwnd || IsWindow(_sharedDirectoriesDialog->hwnd.get()) == FALSE)
    {
        return;
    }

    HRESULT hr = S_OK;
    _sharedDirectoriesDialog->SetRows(CollectSharedDirectoryRows(hr), hr);
}

void FolderWindow::ApplySharedDirectoriesDialogTheme() noexcept
{
    if (! _sharedDirectoriesDialog || ! _sharedDirectoriesDialog->hwnd || IsWindow(_sharedDirectoriesDialog->hwnd.get()) == FALSE)
    {
        return;
    }

    _sharedDirectoriesDialog->theme = _theme;
    _sharedDirectoriesDialog->ApplyTheme();
}

void FolderWindow::RequestCloseSharedDirectoriesDialog() noexcept
{
    if (_hWnd && IsWindow(_hWnd.get()) != FALSE && PostMessageW(_hWnd.get(), WndMsg::kFolderWindowCloseSharedDirectoriesDialog, 0, 0) != 0)
    {
        return;
    }

    CloseSharedDirectoriesDialog();
}

void FolderWindow::CloseSharedDirectoriesDialog() noexcept
{
    if (! _sharedDirectoriesDialog)
    {
        return;
    }

    if (_sharedDirectoriesDialog->hwnd && ! _sharedDirectoriesDialog->destroyed && IsWindow(_sharedDirectoriesDialog->hwnd.get()) != FALSE)
    {
        _sharedDirectoriesDialog->hwnd.reset();
    }
    _sharedDirectoriesDialog.reset();
}

bool FolderWindow::OpenSharedDirectoriesDialogSelection() noexcept
{
    if (! _sharedDirectoriesDialog)
    {
        return false;
    }

    if (_sharedDirectoriesDialog->selectedIndex == kSharedDirectoriesNoSelection ||
        _sharedDirectoriesDialog->selectedIndex >= _sharedDirectoriesDialog->rows.size())
    {
        return false;
    }

    const SharedDirectoryRow row = _sharedDirectoriesDialog->rows[_sharedDirectoriesDialog->selectedIndex];
    if (! row.openable || row.localPath.empty())
    {
        MessageBeep(MB_ICONWARNING);
        return false;
    }

    const Pane pane = _sharedDirectoriesDialog->pane;
    HRESULT hr      = SetFileSystemPluginForPane(pane, L"builtin/file-system");
    if (FAILED(hr))
    {
        std::wstring title   = LoadStringResource(nullptr, IDS_CAPTION_WARNING);
        std::wstring message = LoadStringOrFallback(IDS_SHARED_DIRECTORIES_OPEN_FAILED, L"The shared directory path could not be opened.");
        ShowPaneAlertOverlay(
            pane, FolderView::ErrorOverlayKind::Operation, FolderView::OverlaySeverity::Warning, std::move(title), std::move(message), hr, true, false);
        return false;
    }

    std::error_code ec;
    if (! std::filesystem::is_directory(std::filesystem::path(row.localPath), ec))
    {
        hr                   = HRESULT_FROM_WIN32(ec.value() == 0 ? ERROR_PATH_NOT_FOUND : static_cast<DWORD>(ec.value()));
        std::wstring title   = LoadStringResource(nullptr, IDS_CAPTION_WARNING);
        std::wstring message = LoadStringOrFallback(IDS_SHARED_DIRECTORIES_OPEN_FAILED, L"The shared directory path could not be opened.");
        ShowPaneAlertOverlay(
            pane, FolderView::ErrorOverlayKind::Operation, FolderView::OverlaySeverity::Warning, std::move(title), std::move(message), hr, true, false);
        return false;
    }

    SetActivePane(pane);
    SetFolderPath(pane, std::filesystem::path(row.localPath));
    RequestCloseSharedDirectoriesDialog();
    return true;
}

void FolderWindow::OpenSharedDirectoriesManagement() noexcept
{
    HWND ownerWindow = _hWnd ? GetAncestor(_hWnd.get(), GA_ROOT) : nullptr;
    if (! ownerWindow)
    {
        ownerWindow = _hWnd.get();
    }

    HINSTANCE result = ShellExecuteW(ownerWindow, L"open", L"fsmgmt.msc", nullptr, nullptr, SW_SHOWNORMAL);
    if (reinterpret_cast<INT_PTR>(result) > 32)
    {
        return;
    }

    const Pane pane          = _sharedDirectoriesDialog ? _sharedDirectoriesDialog->pane : _activePane;
    std::wstring title       = LoadStringResource(nullptr, IDS_CAPTION_WARNING);
    std::wstring message     = LoadStringOrFallback(IDS_SHARED_DIRECTORIES_MANAGE_FAILED, L"Windows shared folders management could not be opened.");
    const INT_PTR shellError = reinterpret_cast<INT_PTR>(result);
    const HRESULT resultHr   = shellError == 0 ? E_OUTOFMEMORY : HRESULT_FROM_WIN32(static_cast<DWORD>(shellError));
    ShowPaneAlertOverlay(
        pane, FolderView::ErrorOverlayKind::Operation, FolderView::OverlaySeverity::Warning, std::move(title), std::move(message), resultHr, true, false);
}

void FolderWindow::CommandSharedDirectories(Pane pane)
{
    Debug::Perf::Scope perf(L"sharedDirectories.open_us");
    SetActivePane(pane);

    if (_sharedDirectoriesDialog && _sharedDirectoriesDialog->hwnd && IsWindow(_sharedDirectoriesDialog->hwnd.get()) != FALSE)
    {
        _sharedDirectoriesDialog->pane  = pane;
        _sharedDirectoriesDialog->theme = _theme;
        _sharedDirectoriesDialog->ApplyTheme();
        RefreshSharedDirectoriesDialogRows();
        ShowWindow(_sharedDirectoriesDialog->hwnd.get(), SW_SHOWNORMAL);
        SetForegroundWindow(_sharedDirectoriesDialog->hwnd.get());
        perf.SetValue0(static_cast<uint64_t>(_sharedDirectoriesDialog->rows.size()));
        perf.SetHr(_sharedDirectoriesDialog->lastError);
        return;
    }

    auto state           = std::make_unique<SharedDirectoriesDialogState>();
    state->owner         = this;
    state->pane          = pane;
    state->theme         = _theme;
    state->rows          = CollectSharedDirectoryRows(state->lastError);
    state->selectedIndex = state->rows.empty() ? kSharedDirectoriesNoSelection : 0u;

    HWND ownerWindow = _hWnd ? GetAncestor(_hWnd.get(), GA_ROOT) : nullptr;
    if (! ownerWindow)
    {
        ownerWindow = _hWnd.get();
    }

    const HRESULT classHr = SharedDirectoriesDialogState::EnsureWindowClass();
    if (FAILED(classHr))
    {
        perf.SetHr(classHr);
        Debug::Warning(L"FolderWindow::CommandSharedDirectories: failed to register Shared Directories window class (hr=0x{:08X}).",
                       static_cast<unsigned long>(classHr));
        return;
    }

    const DWORD style   = WS_CAPTION | WS_SYSMENU | WS_POPUP | WS_THICKFRAME | WS_CLIPCHILDREN | WS_CLIPSIBLINGS;
    const DWORD exStyle = WS_EX_DLGMODALFRAME;
    const UINT dpi      = ownerWindow && IsWindow(ownerWindow) != FALSE ? GetDpiForWindow(ownerWindow) : GetDpiForSystem();
    RECT bounds{0, 0, ScalePanePromptForDpi(dpi, 720), ScalePanePromptForDpi(dpi, 460)};
    if (AdjustWindowRectExForDpi(&bounds, style, FALSE, exStyle, dpi) == FALSE)
    {
        const HRESULT hr = HRESULT_FROM_WIN32(GetLastError());
        perf.SetHr(hr);
        Debug::Warning(L"FolderWindow::CommandSharedDirectories: failed to calculate Shared Directories window bounds (hr=0x{:08X}).",
                       static_cast<unsigned long>(hr));
        return;
    }

    const std::wstring caption = LoadStringOrFallback(IDS_SHARED_DIRECTORIES_CAPTION, L"Shared Directories");
    HWND dialog                = CreateWindowExW(exStyle,
                                                 kSharedDirectoriesWindowClassName,
                                                 caption.c_str(),
                                                 style,
                                                 CW_USEDEFAULT,
                                                 CW_USEDEFAULT,
                                                 bounds.right - bounds.left,
                                                 bounds.bottom - bounds.top,
                                                 ownerWindow,
                                                 nullptr,
                                                 GetModuleHandleW(nullptr),
                                                 state.get());
    if (! dialog)
    {
        const HRESULT hr = HRESULT_FROM_WIN32(GetLastError());
        perf.SetHr(hr);
        Debug::Warning(L"FolderWindow::CommandSharedDirectories: failed to create Shared Directories dialog (hr=0x{:08X}).", static_cast<unsigned long>(hr));
        return;
    }

    if (! state->hwnd)
    {
        state->hwnd = wil::unique_hwnd(dialog);
    }
    perf.SetValue0(static_cast<uint64_t>(state->rows.size()));
    perf.SetHr(state->lastError);
    _sharedDirectoriesDialog = std::move(state);
    static_cast<void>(Common::WindowSizing::CenterExistingWindowOnOwner(dialog, ownerWindow));
    static_cast<void>(_sharedDirectoriesDialog->dxHost.PrimeForShow());
    ShowWindow(dialog, SW_SHOWNORMAL);
    UpdateWindow(dialog);
    SetForegroundWindow(dialog);
}

void FolderWindow::CommandSelectionRestore(Pane pane)
{
    SetActivePane(pane);
    PaneState& state = pane == Pane::Left ? _leftPane : _rightPane;

    if (! HasSavedSelection())
    {
        MessageBeep(MB_ICONWARNING);
        std::wstring title   = LoadStringResource(nullptr, IDS_CAPTION_WARNING);
        std::wstring message = LoadStringResource(nullptr, IDS_MSG_SELECTION_RESTORE_NO_SAVED);
        ShowPaneAlertOverlay(
            pane, FolderView::ErrorOverlayKind::Operation, FolderView::OverlaySeverity::Warning, std::move(title), std::move(message), S_OK, true, false);
        return;
    }

    const SavedSelection& saved = _savedSelection.value();

    std::unordered_set<std::wstring_view> remaining;
    remaining.reserve(saved.displayNames.size());
    for (const auto& name : saved.displayNames)
    {
        if (! name.empty())
        {
            remaining.emplace(name);
        }
    }

    SetPaneSelectionByDisplayNamePredicate(pane,
                                           [&](std::wstring_view displayName) noexcept
    {
        const auto it = remaining.find(displayName);
        if (it == remaining.end())
        {
            return false;
        }
        remaining.erase(it);
        return true;
    },
                                           true /* clearExistingSelection */);

    if (! remaining.empty())
    {
        struct WStringViewNoCaseLess final
        {
            bool operator()(std::wstring_view left, std::wstring_view right) const noexcept
            {
                return OrdinalString::Compare(left, right, true) < 0;
            }
        };

        std::map<std::wstring_view, std::vector<std::wstring_view>, WStringViewNoCaseLess> remainingNoCase;
        for (const auto& name : remaining)
        {
            remainingNoCase[name].push_back(name);
        }

        SetPaneSelectionByDisplayNamePredicate(pane,
                                               [&](std::wstring_view displayName) noexcept
        {
            const auto it = remainingNoCase.find(displayName);
            if (it == remainingNoCase.end() || it->second.empty())
            {
                return false;
            }

            const std::wstring_view matched = it->second.back();
            it->second.pop_back();
            remaining.erase(matched);
            if (it->second.empty())
            {
                remainingNoCase.erase(it);
            }
            return true;
        },
                                               false /* clearExistingSelection */);
    }

    if (remaining.empty())
    {
        return;
    }

    std::wstring missingLines;
    for (const auto& name : saved.displayNames)
    {
        if (name.empty())
        {
            continue;
        }

        if (remaining.find(std::wstring_view(name)) == remaining.end())
        {
            continue;
        }

        missingLines.append(L"- ");
        missingLines.append(name);
        missingLines.append(L"\r\n");
    }

    std::wstring filterNote;
    if (state.folderView.IsNameFilterActive())
    {
        const std::wstring noteText = LoadStringResource(nullptr, IDS_MSG_SELECTION_RESTORE_FILTER_NOTE);
        if (! noteText.empty())
        {
            filterNote.reserve(4u + noteText.size());
            filterNote.append(L"\r\n\r\n");
            filterNote.append(noteText);
        }
    }

    std::wstring title = LoadStringResource(nullptr, IDS_CMD_SELECTION_RESTORE);
    if (title.empty())
    {
        title = LoadStringResource(nullptr, IDS_OVERLAY_TITLE_INFORMATION);
    }

    std::wstring message = FormatStringResource(nullptr, IDS_FMT_SELECTION_RESTORE_INCOMPLETE, missingLines, filterNote);
    ShowPaneAlertOverlay(
        pane, FolderView::ErrorOverlayKind::Operation, FolderView::OverlaySeverity::Information, std::move(title), std::move(message), S_OK, true, false);
}

void FolderWindow::CommandSelectionSelectSameExtension(Pane pane)
{
    SetActivePane(pane);
    PaneState& state = pane == Pane::Left ? _leftPane : _rightPane;
    state.folderView.SelectSameExtension();
}

void FolderWindow::CommandSelectionSelectSameName(Pane pane)
{
    SetActivePane(pane);
    PaneState& state = pane == Pane::Left ? _leftPane : _rightPane;
    state.folderView.SelectSameName();
}

void FolderWindow::CommandSelectionUnselectSameExtension(Pane pane)
{
    SetActivePane(pane);
    PaneState& state = pane == Pane::Left ? _leftPane : _rightPane;
    state.folderView.UnselectSameExtension();
}

void FolderWindow::CommandSelectionUnselectSameName(Pane pane)
{
    SetActivePane(pane);
    PaneState& state = pane == Pane::Left ? _leftPane : _rightPane;
    state.folderView.UnselectSameName();
}

void FolderWindow::CommandSelectionHideSelectedNames(Pane pane)
{
    SetActivePane(pane);
    PaneState& state = pane == Pane::Left ? _leftPane : _rightPane;
    state.folderView.HideSelectedNames();
}

void FolderWindow::CommandSelectionHideUnselectedNames(Pane pane)
{
    SetActivePane(pane);
    PaneState& state = pane == Pane::Left ? _leftPane : _rightPane;
    state.folderView.HideUnselectedNames();
}

void FolderWindow::CommandSelectionShowHiddenNames(Pane pane)
{
    SetActivePane(pane);
    PaneState& state = pane == Pane::Left ? _leftPane : _rightPane;
    state.folderView.ShowHiddenNames();
}

bool FolderWindow::CanShowHiddenNames(Pane pane) const noexcept
{
    const PaneState& state = pane == Pane::Left ? _leftPane : _rightPane;
    return state.folderView.HasHiddenNames();
}

void FolderWindow::CommandChangeAttributes(Pane pane)
{
    const auto startedAt = std::chrono::steady_clock::now();
    SetActivePane(pane);
    PaneState& state = pane == Pane::Left ? _leftPane : _rightPane;

#ifdef ENABLE_TESTS
    _debugLastChangeAttributesReport.reset();
#endif

    if (! state.fileSystem)
    {
        ShowChangeAttributesOverlay(*this, pane, FolderView::OverlaySeverity::Warning, IDS_MSG_CHANGE_ATTRIBUTES_UNSUPPORTED);
        Debug::Perf::Emit(L"fileattrs.change_attributes_us", L"unsupported", Debug::Perf::ElapsedUs(startedAt), 0u, 0u, E_NOINTERFACE);
        return;
    }

    const std::vector<std::filesystem::path> paths = state.folderView.GetSelectedOrFocusedPaths();
    if (paths.empty())
    {
        ShowChangeAttributesOverlay(*this, pane, FolderView::OverlaySeverity::Warning, IDS_MSG_CHANGE_ATTRIBUTES_SELECTION_REQUIRED);
        Debug::Perf::Emit(L"fileattrs.change_attributes_us", L"empty-selection", Debug::Perf::ElapsedUs(startedAt), 0u, 0u, S_FALSE);
        return;
    }

    wil::com_ptr<IFileSystemIO> fileIo;
    const HRESULT fileIoHr = state.fileSystem->QueryInterface(__uuidof(IFileSystemIO), fileIo.put_void());
    if (FAILED(fileIoHr) || ! fileIo)
    {
        ShowChangeAttributesOverlay(*this, pane, FolderView::OverlaySeverity::Warning, IDS_MSG_CHANGE_ATTRIBUTES_UNSUPPORTED, fileIoHr);
        Debug::Perf::Emit(L"fileattrs.change_attributes_us", L"unsupported", Debug::Perf::ElapsedUs(startedAt), 0u, 0u, fileIoHr);
        return;
    }

    const bool allowSubdirectories         = ChangeAttributesSelectionContainsDirectory(fileIo.get(), paths);
    ChangeAttributesOptions initialOptions = BuildInitialChangeAttributesOptions(fileIo.get(), paths);
    initialOptions.includeSubdirectories   = false;

    std::optional<ChangeAttributesOptions> optionsOpt;
#ifdef ENABLE_TESTS
    if (_debugNextChangeAttributesOptions.has_value())
    {
        optionsOpt = _debugNextChangeAttributesOptions.value();
    }
    else
#endif
    {
        HWND ownerWindow = _hWnd ? GetAncestor(_hWnd.get(), GA_ROOT) : nullptr;
        if (! ownerWindow)
        {
            ownerWindow = _hWnd.get();
        }
        optionsOpt = PromptForChangeAttributes(ownerWindow, _theme, initialOptions, allowSubdirectories);
    }

    if (! optionsOpt.has_value())
    {
        Debug::Perf::Emit(L"fileattrs.change_attributes_us", L"cancelled", Debug::Perf::ElapsedUs(startedAt), 0u, 0u, S_FALSE);
        return;
    }

    const ChangeAttributesOptions options = optionsOpt.value();
    if (IsChangeAttributesNoop(options))
    {
        ShowChangeAttributesOverlay(*this, pane, FolderView::OverlaySeverity::Warning, IDS_MSG_CHANGE_ATTRIBUTES_NO_CHANGES, S_FALSE);
        Debug::Perf::Emit(L"fileattrs.change_attributes_us", L"no-changes", Debug::Perf::ElapsedUs(startedAt), 0u, 0u, S_FALSE);
        return;
    }

    wil::com_ptr<IFileSystemItemStreams> streamOps;
    if (options.removeAlternateDataStreams)
    {
        static_cast<void>(state.fileSystem->QueryInterface(__uuidof(IFileSystemItemStreams), streamOps.put_void()));
    }

    if (options.includeSubdirectories && allowSubdirectories)
    {
        if (state.changeAttributesThread.joinable())
        {
            MessageBeep(MB_ICONWARNING);
            return;
        }

        InformationalTaskUpdate task{};
        task.kind                         = InformationalTaskUpdate::Kind::ChangeAttributes;
        task.title                        = LoadStringResource(nullptr, IDS_CMD_CHANGE_ATTRIBUTES);
        task.changeAttributesCurrentPath  = paths.front();
        task.changeAttributesEnumerating  = true;
        task.changeAttributesPlannedItems = static_cast<uint64_t>(paths.size());
        const uint64_t taskId             = CreateOrUpdateInformationalTask(task);

        const HWND ownerHwnd                 = _hWnd.get();
        wil::com_ptr<IFileSystem> fileSystem = state.fileSystem;
        const std::wstring title             = task.title;

        state.changeAttributesThread = std::jthread([ownerHwnd, pane, fileSystem, paths, options, title, taskId](std::stop_token stopToken) noexcept
        {
            const auto operationStartedAt = std::chrono::steady_clock::now();
            const HRESULT coinitHr        = CoInitializeEx(nullptr, COINIT_MULTITHREADED);
            if (FAILED(coinitHr))
            {
                Debug::Error(L"ChangeAttributes task: CoInitializeEx(COINIT_MULTITHREADED) failed: 0x{:08X}", coinitHr);
                FAIL_FAST_IF_FAILED(coinitHr);
            }
            [[maybe_unused]] const wil::unique_couninitialize_call coUninit;

            ChangeAttributesReport report{};
            report.progressTaskId = taskId;
            bool refreshNeeded    = false;
            HRESULT operationHr   = S_OK;

            ChangeAttributesProgressState progress{};
            progress.hwnd         = ownerHwnd;
            progress.pane         = pane;
            progress.title        = title;
            progress.taskId       = taskId;
            progress.plannedItems = static_cast<uint64_t>(paths.size());
            progress.PostTaskUpdate(false, S_OK);

            wil::com_ptr<IFileSystemIO> workerFileIo;
            HRESULT hr = fileSystem ? fileSystem->QueryInterface(__uuidof(IFileSystemIO), workerFileIo.put_void()) : E_POINTER;
            if (FAILED(hr) || ! workerFileIo)
            {
                operationHr = FAILED(hr) ? hr : E_POINTER;
                RecordChangeAttributesFailure(report, operationHr);
            }

            wil::com_ptr<IFileSystemItemStreams> workerStreamOps;
            if (SUCCEEDED(operationHr) && options.removeAlternateDataStreams)
            {
                static_cast<void>(fileSystem->QueryInterface(__uuidof(IFileSystemItemStreams), workerStreamOps.put_void()));
            }

            std::vector<std::filesystem::path> workItems;
            if (SUCCEEDED(operationHr))
            {
                hr = CollectRecursiveChangeAttributesPaths(fileSystem.get(), workerFileIo.get(), paths, stopToken, progress, report, workItems);
                if (FAILED(hr))
                {
                    operationHr = hr;
                    RecordChangeAttributesFailure(report, hr);
                }
            }

            progress.enumerating  = false;
            progress.applying     = true;
            progress.plannedItems = static_cast<uint64_t>(workItems.size());
            progress.MaybePostTaskUpdate();

            if (SUCCEEDED(operationHr))
            {
                for (const std::filesystem::path& path : workItems)
                {
                    if (stopToken.stop_requested())
                    {
                        operationHr = HRESULT_FROM_WIN32(ERROR_CANCELLED);
                        break;
                    }

                    progress.currentPath = path;
                    ApplyChangeAttributesToPath(workerFileIo.get(), workerStreamOps.get(), path, options, report, refreshNeeded);
                    ++progress.completedItems;
                    progress.MaybePostTaskUpdate();
                }
            }

            report.summary = BuildChangeAttributesReportSummary(report);
            progress.PostTaskUpdate(true, operationHr, report.summary);

            if (ownerHwnd && IsWindow(ownerHwnd) != FALSE)
            {
                auto completed           = std::make_unique<ChangeAttributesCompletedPayload>();
                completed->pane          = pane;
                completed->hr            = operationHr;
                completed->report        = std::move(report);
                completed->refreshNeeded = refreshNeeded;
                static_cast<void>(PostMessagePayload(ownerHwnd, WndMsg::kChangeAttributesCompleted, 0, std::move(completed)));
            }

            Debug::Perf::Emit(
                L"fileattrs.change_attributes_us", L"recursive", Debug::Perf::ElapsedUs(operationStartedAt), progress.completedItems, 0u, operationHr);
        });
        return;
    }

    ChangeAttributesReport report{};
    bool refreshNeeded = false;
    for (const std::filesystem::path& path : paths)
    {
        ApplyChangeAttributesToPath(fileIo.get(), streamOps.get(), path, options, report, refreshNeeded);
    }
    report.summary = BuildChangeAttributesReportSummary(report);
#ifdef ENABLE_TESTS
    _debugLastChangeAttributesReport = report;
#endif

    if (refreshNeeded)
    {
        state.folderView.ForceRefresh();
    }

    ShowChangeAttributesReportOverlay(*this, pane, report);
    Debug::Perf::Emit(L"fileattrs.change_attributes_us",
                      L"",
                      Debug::Perf::ElapsedUs(startedAt),
                      report.itemsProcessed,
                      report.streamsRemoved,
                      report.failures == 0u ? S_OK : report.firstFailure);
}

void FolderWindow::CommandSelectionGoToPreviousSelectedName(Pane pane)
{
    SetActivePane(pane);
    PaneState& state = pane == Pane::Left ? _leftPane : _rightPane;
    static_cast<void>(state.folderView.GoToPreviousSelectedName());
}

void FolderWindow::CommandSelectionGoToNextSelectedName(Pane pane)
{
    SetActivePane(pane);
    PaneState& state = pane == Pane::Left ? _leftPane : _rightPane;
    static_cast<void>(state.folderView.GoToNextSelectedName());
}

void FolderWindow::CommandChangeCase(Pane pane)
{
    SetActivePane(pane);
    PaneState& state = pane == Pane::Left ? _leftPane : _rightPane;

    if (! state.fileSystem)
    {
        return;
    }

    if (state.changeCaseThread.joinable())
    {
        MessageBeep(MB_ICONWARNING);
        return;
    }

    HWND ownerWindow = _hWnd ? GetAncestor(_hWnd.get(), GA_ROOT) : nullptr;
    if (! ownerWindow)
    {
        ownerWindow = _hWnd.get();
    }

    const std::optional<ChangeCase::Options> dialogResult = PromptForChangeCase(ownerWindow, _theme, true);
    if (! dialogResult.has_value())
    {
        return;
    }

    const std::vector<std::filesystem::path> paths = state.folderView.GetSelectedOrFocusedPaths();
    if (paths.empty())
    {
        MessageBeep(MB_ICONWARNING);
        return;
    }

    const HWND ownerHwnd                 = _hWnd.get();
    wil::com_ptr<IFileSystem> fileSystem = state.fileSystem;
    const ChangeCase::Options options    = dialogResult.value();
    const std::wstring title             = LoadStringResource(nullptr, IDS_CMD_CHANGE_CASE);

    state.changeCaseThread = std::jthread([ownerHwnd, pane, fileSystem, paths, options, title](std::stop_token stopToken) noexcept
    {
        const HRESULT coinitHr = CoInitializeEx(nullptr, COINIT_MULTITHREADED);
        if (FAILED(coinitHr))
        {
            Debug::Error(L"ChangeCase task: CoInitializeEx(COINIT_MULTITHREADED) failed: 0x{:08X}", coinitHr);
            FAIL_FAST_IF_FAILED(coinitHr);
        }
        [[maybe_unused]] const wil::unique_couninitialize_call coUninit;

        struct ProgressState final
        {
            HWND hwnd               = nullptr;
            FolderWindow::Pane pane = FolderWindow::Pane::Left;
            std::wstring title;
            ULONGLONG startTick      = 0;
            ULONGLONG lastPostedTick = 0;
            uint64_t infoTaskId      = 0;
            ChangeCase::ProgressUpdate last{};

            void PostTaskUpdate(bool finished, HRESULT hr) noexcept
            {
                if (! hwnd || IsWindow(hwnd) == FALSE || infoTaskId == 0)
                {
                    return;
                }

                FolderWindow::InformationalTaskUpdate info{};
                info.kind                       = FolderWindow::InformationalTaskUpdate::Kind::ChangeCase;
                info.taskId                     = infoTaskId;
                info.title                      = title;
                info.changeCaseCurrentPath      = last.currentPath;
                info.changeCaseScannedFolders   = last.scannedFolders;
                info.changeCaseScannedEntries   = last.scannedEntries;
                info.changeCasePlannedRenames   = last.plannedRenames;
                info.changeCaseCompletedRenames = last.completedRenames;
                info.changeCaseEnumerating      = ! finished && last.phase == ChangeCase::ProgressUpdate::Phase::Enumerating;
                info.changeCaseRenaming         = ! finished && last.phase == ChangeCase::ProgressUpdate::Phase::Renaming;
                info.finished                   = finished;
                info.resultHr                   = hr;

                auto payload    = std::make_unique<ChangeCaseTaskPayload>();
                payload->update = std::move(info);
                static_cast<void>(PostMessagePayload(hwnd, WndMsg::kChangeCaseTaskUpdate, 0, std::move(payload)));
            }

            void EnsureTaskVisibleAfterThreshold() noexcept
            {
                if (! hwnd || IsWindow(hwnd) == FALSE || infoTaskId != 0)
                {
                    return;
                }

                const ULONGLONG nowTick = GetTickCount64();
                if (startTick == 0 || nowTick < startTick || (nowTick - startTick) < 700ull)
                {
                    return;
                }

                FolderWindow::InformationalTaskUpdate info{};
                info.kind                       = FolderWindow::InformationalTaskUpdate::Kind::ChangeCase;
                info.title                      = title;
                info.changeCaseCurrentPath      = last.currentPath;
                info.changeCaseScannedFolders   = last.scannedFolders;
                info.changeCaseScannedEntries   = last.scannedEntries;
                info.changeCasePlannedRenames   = last.plannedRenames;
                info.changeCaseCompletedRenames = last.completedRenames;
                info.changeCaseEnumerating      = last.phase == ChangeCase::ProgressUpdate::Phase::Enumerating;
                info.changeCaseRenaming         = last.phase == ChangeCase::ProgressUpdate::Phase::Renaming;

                auto payload    = std::make_unique<ChangeCaseTaskPayload>();
                payload->update = std::move(info);

                ChangeCaseTaskPayload* raw = payload.release();
                DWORD_PTR result           = 0;
                const LRESULT sendOk =
                    SendMessageTimeoutW(hwnd, WndMsg::kChangeCaseTaskUpdate, 0, reinterpret_cast<LPARAM>(raw), SMTO_ABORTIFHUNG | SMTO_BLOCK, 1000, &result);
                if (sendOk == 0)
                {
                    delete raw;
                    return;
                }

                infoTaskId     = static_cast<uint64_t>(result);
                lastPostedTick = nowTick;
            }
        };

        ProgressState progressState{};
        progressState.hwnd      = ownerHwnd;
        progressState.pane      = pane;
        progressState.title     = title;
        progressState.startTick = GetTickCount64();

        const auto onProgress = [](const ChangeCase::ProgressUpdate& update, void* cookie) noexcept
        {
            auto* state = static_cast<ProgressState*>(cookie);
            if (! state)
            {
                return;
            }

            state->last = update;
            state->EnsureTaskVisibleAfterThreshold();

            if (state->infoTaskId != 0)
            {
                const ULONGLONG nowTick = GetTickCount64();
                if (state->lastPostedTick != 0 && nowTick >= state->lastPostedTick && (nowTick - state->lastPostedTick) < 100ull)
                {
                    return;
                }

                state->lastPostedTick = nowTick;
                state->PostTaskUpdate(false, S_OK);
            }
        };

        const HRESULT operationHr = ChangeCase::ApplyToPaths(*fileSystem, paths, options, stopToken, onProgress, &progressState);

        progressState.EnsureTaskVisibleAfterThreshold();
        progressState.PostTaskUpdate(true, operationHr);

        if (ownerHwnd && IsWindow(ownerHwnd) != FALSE)
        {
            auto completed  = std::make_unique<ChangeCaseCompletedPayload>();
            completed->pane = pane;
            completed->hr   = operationHr;
            static_cast<void>(PostMessagePayload(ownerHwnd, WndMsg::kChangeCaseCompleted, 0, std::move(completed)));
        }
    });
}
