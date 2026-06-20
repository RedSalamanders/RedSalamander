#pragma once

#include <cstdint>
#include <optional>
#include <string>
#include <string_view>

enum class FileSystemPathComponentComparison : uint8_t
{
    OrdinalIgnoreCase,
    OrdinalCaseSensitive,
};

enum class FileSystemPathCaseOnlyRename : uint8_t
{
    Supported,
    NoOp,
    Unsupported,
    NotApplicable,
};

struct FileSystemPathIdentity final
{
    bool pathTextStableIdentity = true;
    FileSystemPathComponentComparison componentComparison = FileSystemPathComponentComparison::OrdinalIgnoreCase;
    wchar_t preferredSeparator = L'\\';
    std::wstring acceptedSeparators = L"\\/";
    bool casePreserving = true;
    FileSystemPathCaseOnlyRename caseOnlyRename = FileSystemPathCaseOnlyRename::Supported;

    [[nodiscard]] static FileSystemPathIdentity OrdinalIgnoreCaseForLocalFileSystem() noexcept;
};

[[nodiscard]] std::optional<FileSystemPathIdentity> TryParseFileSystemPathIdentityContract(std::string_view jsonUtf8,
                                                                                           std::wstring_view pluginId) noexcept;
[[nodiscard]] std::optional<FileSystemPathIdentity> TryParseFileSystemPathIdentity(std::string_view jsonUtf8,
                                                                                   std::wstring_view pluginId) noexcept;
[[nodiscard]] std::optional<FileSystemPathIdentity> TryParseFileSystemRenamePathIdentity(std::string_view jsonUtf8,
                                                                                         std::wstring_view pluginId) noexcept;
[[nodiscard]] bool EquivalentComponent(const FileSystemPathIdentity& identity, std::wstring_view lhs, std::wstring_view rhs) noexcept;
[[nodiscard]] bool EquivalentPath(const FileSystemPathIdentity& identity, std::wstring_view lhs, std::wstring_view rhs) noexcept;
[[nodiscard]] std::optional<std::wstring> TryMakeComponentKey(const FileSystemPathIdentity& identity,
                                                              std::wstring_view component) noexcept;
[[nodiscard]] std::optional<std::wstring> TryMakePathKey(const FileSystemPathIdentity& identity, std::wstring_view path) noexcept;
