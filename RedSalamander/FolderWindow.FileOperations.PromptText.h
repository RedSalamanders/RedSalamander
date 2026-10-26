#pragma once

#include "PlugInterfaces/FileSystem.h"

#include <cstdint>
#include <string>
#include <string_view>

// I35 R8.3: the text of a published conflict or consent prompt, shared by the File Operations task card
// and the decision mirror over a pane, so both renderers ask the same question in the same words. The
// inputs are the published prompt's facts; nothing here selects, orders or places an action (the engine
// policy owns that). Views must outlive the call.
namespace FileOperations
{
struct ConflictPromptTextFacts final
{
    FileSystemOperation operation   = FILESYSTEM_COPY;
    uint8_t bucket                  = 0u; // ConflictClass
    HRESULT status                  = S_OK;
    unsigned int attemptCount       = 0u;
    uint8_t overlapProblem          = 0u; // SameHostOverlapProblem
    uint64_t overlapTaskId          = 0u;
    uint8_t repeatUnavailableReason = 0u; // ConflictRepeatUnavailableReason
    bool readOnlyDeleteOffered      = false;
    bool sourceIsDirectory          = false;
    bool skipAllOffered             = false;
    std::wstring_view sourcePath;
    std::wstring_view destinationPath;
    std::wstring_view sourceName;
    std::wstring_view sourceParentPath;
    std::wstring_view destinationName;
    std::wstring_view destinationParentPath;
    std::wstring_view consentDetail;
    // I35 R7.3/R7.6: how the location deletes, and the leftover warning; closes a consent question as its own paragraph.
    std::wstring_view consentNote;
    bool factItemCountKnown = false;
    uint64_t factItemCount  = 0u;
    bool factBytesKnown     = false;
    uint64_t factBytes      = 0u;
};

// The one leaf both roles name, used as a heading; empty unless a name collision names the same admitted leaf.
[[nodiscard]] std::wstring_view CommonConflictPromptName(const ConflictPromptTextFacts& facts) noexcept;
// The question the prompt asks.
[[nodiscard]] std::wstring BuildConflictPromptQuestion(const ConflictPromptTextFacts& facts);
// The item and byte counts a consent names; empty when none is known.
[[nodiscard]] std::wstring BuildConflictPromptFactText(const ConflictPromptTextFacts& facts);
// What a screen reader is told once when the decision is published: the waiting state, the shared name
// and the question (R8.12: the card and the decision mirror announce the same text, never both).
[[nodiscard]] std::wstring BuildConflictPromptAnnouncement(const ConflictPromptTextFacts& facts);
// The operation's card title (Copy, Move, Delete, Rename).
[[nodiscard]] UINT ConflictPromptTitleStringId(FileSystemOperation operation) noexcept;
} // namespace FileOperations
