#pragma once

#include "framework.h"
#include "FileSystemPathIdentity.h"

#include <chrono>
#include <cstdint>
#include <filesystem>
#include <optional>
#include <span>
#include <string>
#include <string_view>
#include <vector>

namespace BatchRename
{
enum class Mode : uint8_t
{
    Rules,
    Manual,
};

enum class CaseTransform : uint8_t
{
    None,
    Lower,
    Upper,
    Mixed,
};

enum class IssueSeverity : uint8_t
{
    Warning,
    Error,
};

struct Target final
{
    std::filesystem::path sourcePath;
    std::filesystem::path relativeFolder;
    bool isDirectory  = false;
    bool metadataUnknown = false;
    uint64_t sizeBytes = 0;
    std::optional<std::chrono::sys_seconds> lastWriteTime;
    std::optional<std::chrono::sys_seconds> createdTime;
};

struct Rules final
{
    Mode mode = Mode::Rules;

    std::wstring nameTemplate = L"{name}";
    std::wstring searchFor;
    std::wstring replaceWith;
    bool regexEnabled      = false;
    bool caseSensitive     = true;
    bool wholeWords        = false;
    bool replaceOnce       = false;
    bool excludeExtension  = false;
    std::wstring flattenSeparator = L" - ";

    CaseTransform fileNameCaseStyle  = CaseTransform::None;
    CaseTransform extensionCaseStyle = CaseTransform::None;

    std::vector<std::wstring> manualNames;
};

struct Issue final
{
    IssueSeverity severity = IssueSeverity::Error;
    std::wstring message;
};

struct PreviewRow final
{
    uint64_t rowId = 0;
    std::filesystem::path sourcePath;
    std::wstring originalName;
    std::wstring newName;
    bool isDirectory  = false;
    bool metadataUnknown = false;
    uint64_t sizeBytes = 0;
    std::optional<std::chrono::sys_seconds> lastWriteTime;
    std::optional<std::chrono::sys_seconds> createdTime;
    std::vector<Issue> issues;
};

struct Stats final
{
    size_t totalRows     = 0;
    size_t changedRows   = 0;
    size_t unchangedRows = 0;
    size_t errorRows     = 0;
    size_t warningRows   = 0;
};

struct Plan final
{
    std::vector<PreviewRow> rows;
    Stats stats;
};

void AddIssue(PreviewRow& row, IssueSeverity severity, std::wstring message);
[[nodiscard]] bool HasIssueSeverity(const PreviewRow& row, IssueSeverity severity) noexcept;
[[nodiscard]] Stats RecomputeStats(std::span<const PreviewRow> rows) noexcept;
void RecomputeStats(Plan& plan) noexcept;
[[nodiscard]] std::chrono::local_seconds ToLocalWallClock(std::chrono::sys_seconds timestamp) noexcept;
[[nodiscard]] std::wstring FormatTimestamp(std::chrono::sys_seconds timestamp, std::wstring_view format);
[[nodiscard]] std::wstring FormatDateText(std::chrono::sys_seconds timestamp);
[[nodiscard]] std::wstring FormatTimeText(std::chrono::sys_seconds timestamp);
[[nodiscard]] Plan BuildPlan(const std::vector<Target>& targets, const Rules& rules) noexcept;
[[nodiscard]] Plan BuildPlan(const std::vector<Target>& targets, const Rules& rules, const FileSystemPathIdentity& pathIdentity) noexcept;
} // namespace BatchRename
