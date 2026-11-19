#pragma once

#include <cstddef>
#include <stop_token>
#include <string>
#include <string_view>
#include <vector>

namespace TerminalCommandSurfaceModel
{
struct FindMatch final
{
    size_t lineNumber  = 0u;
    size_t matchOffset = 0u;
    std::wstring text;
};

struct FindResult final
{
    std::vector<FindMatch> rows;
    size_t totalMatches = 0u;
    size_t scannedLines = 0u;
    bool cancelled      = false;
};

struct Suggestion final
{
    std::wstring text;
    bool currentSession = false;
};

struct SuggestionResult final
{
    std::vector<Suggestion> rows;
    size_t sourceEntries = 0u;
    size_t sourceBytes   = 0u;
    bool cancelled       = false;
};

[[nodiscard]] FindResult Find(std::wstring_view snapshot, std::wstring_view query, size_t maximumRows, std::stop_token stopToken = {});
[[nodiscard]] std::vector<std::wstring> LoadPowerShellHistory(std::wstring_view trustedPath,
                                                              size_t maximumBytes,
                                                              size_t maximumEntries,
                                                              std::stop_token stopToken = {});
[[nodiscard]] SuggestionResult FilterSuggestions(const std::vector<std::wstring>& persistedHistory,
                                                 const std::vector<std::wstring>& currentSessionHistory,
                                                 std::wstring_view query,
                                                 size_t maximumRows,
                                                 std::stop_token stopToken = {});
[[nodiscard]] bool IsSafeSuggestion(std::wstring_view text) noexcept;
} // namespace TerminalCommandSurfaceModel
