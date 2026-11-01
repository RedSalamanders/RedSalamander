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

// Relation between two keys produced by TryMakePathKey under the same path contract.
enum class FileSystemPathKeyRelation : uint8_t
{
    Disjoint,
    Equal,
    LeftAncestor,
    RightAncestor,
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
// Orders paths by their components under the supplied comparison and separator profile, with an
// ancestor before every descendant. This keeps each component subtree contiguous even when
// punctuation in a sibling sorts before the separator.
[[nodiscard]] int CompareFileSystemPathsByComponents(const FileSystemPathIdentity& identity, std::wstring_view lhs, std::wstring_view rhs) noexcept;
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
// EquivalentComponent/EquivalentPath. Ordinal-ignore-case keys fold UTF-16 units separately,
// preserving surrogate units and validating BMP mappings with CompareStringOrdinal. Whole-string
// Unicode uppercasing is incompatible with this relation. Callers confirm key hits with raw equality.
[[nodiscard]] std::optional<std::wstring> TryMakeComponentKey(const FileSystemPathIdentity& identity, std::wstring_view component) noexcept;
[[nodiscard]] std::optional<std::wstring> TryMakePathKey(const FileSystemPathIdentity& identity, std::wstring_view path) noexcept;
// Compares already-produced canonical keys without allocating. A missing result means the supplied
// identity cannot authorize key comparison; callers must use their exact path fallback.
[[nodiscard]] std::optional<FileSystemPathKeyRelation> CompareFileSystemPathKeys(const FileSystemPathIdentity& identity,
                                                                                 std::wstring_view leftKey,
                                                                                 std::wstring_view rightKey) noexcept;
