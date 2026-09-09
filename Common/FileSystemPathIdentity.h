#pragma once

#include <cstdint>
#include <optional>
#include <string>
#include <string_view>

struct yyjson_val;

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
    bool pathTextStableIdentity                           = true;
    FileSystemPathComponentComparison componentComparison = FileSystemPathComponentComparison::OrdinalIgnoreCase;
    wchar_t preferredSeparator                            = L'\\';
    std::wstring acceptedSeparators                       = L"\\/";
    bool casePreserving                                   = true;
    FileSystemPathCaseOnlyRename caseOnlyRename           = FileSystemPathCaseOnlyRename::Supported;

    [[nodiscard]] static FileSystemPathIdentity OrdinalIgnoreCaseForLocalFileSystem() noexcept;
};

// Diagnostics/tests only. Executable route and path-identity decisions must use
// FileSystemRouteContract and IFileSystemRouteCapabilities.
[[nodiscard]] std::optional<FileSystemPathIdentity> ParseDiagnosticFileSystemPathIdentityContract(std::string_view jsonUtf8,
                                                                                                  std::wstring_view pluginId) noexcept;
[[nodiscard]] std::optional<FileSystemPathIdentity> ParseDiagnosticFileSystemPathIdentityContractFromRoot(yyjson_val* root,
                                                                                                          std::wstring_view pluginId) noexcept;
[[nodiscard]] std::optional<FileSystemPathIdentity> ParseDiagnosticFileSystemPathIdentity(std::string_view jsonUtf8, std::wstring_view pluginId) noexcept;
[[nodiscard]] std::optional<FileSystemPathIdentity> ParseDiagnosticFileSystemRenamePathIdentity(std::string_view jsonUtf8, std::wstring_view pluginId) noexcept;
[[nodiscard]] bool EquivalentComponent(const FileSystemPathIdentity& identity, std::wstring_view lhs, std::wstring_view rhs) noexcept;
[[nodiscard]] bool EquivalentPath(const FileSystemPathIdentity& identity, std::wstring_view lhs, std::wstring_view rhs) noexcept;
[[nodiscard]] bool TryGetFileSystemParentPath(const FileSystemPathIdentity& identity, std::wstring_view path, std::wstring& parentOut) noexcept;
[[nodiscard]] bool TryGetFileSystemLeafName(const FileSystemPathIdentity& identity, std::wstring_view path, std::wstring& leafOut) noexcept;
[[nodiscard]] std::wstring JoinFileSystemPath(const FileSystemPathIdentity& identity, std::wstring_view folder, std::wstring_view leaf);
[[nodiscard]] bool IsStrictDescendantPath(const FileSystemPathIdentity& identity, std::wstring_view prefix, std::wstring_view candidate) noexcept;
[[nodiscard]] bool TryGetFileSystemRelativePath(const FileSystemPathIdentity& identity,
                                                std::wstring_view root,
                                                std::wstring_view candidate,
                                                std::wstring& relativeOut) noexcept;
[[nodiscard]] std::wstring ReplaceFileSystemPathPrefix(const FileSystemPathIdentity& identity,
                                                       std::wstring_view candidate,
                                                       std::wstring_view oldPrefix,
                                                       std::wstring_view newPrefix);
// nullopt means the identity cannot be represented safely as a hash key; callers must fall back to
// EquivalentComponent/EquivalentPath. Ordinal-ignore-case keys use Windows invariant simple uppercase
// folding as an approximation of NTFS $UpCase and reject expanding/otherwise unsafe mappings.
[[nodiscard]] std::optional<std::wstring> TryMakeComponentKey(const FileSystemPathIdentity& identity, std::wstring_view component) noexcept;
[[nodiscard]] std::optional<std::wstring> TryMakePathKey(const FileSystemPathIdentity& identity, std::wstring_view path) noexcept;
