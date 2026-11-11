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
// I62 B3: what is known about one role's item on a name collision.
struct ConflictPromptItemFacts final
{
    bool available        = false;
    bool isDirectory      = false;
    bool sizeKnown        = false;
    uint64_t sizeBytes    = 0u;
    __int64 lastWriteTime = 0; // FILETIME ticks; zero or less is unknown
};

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
    // I62 B3: both items' facts and the engine's comparison of the incoming item with the existing one.
    ConflictPromptItemFacts sourceItem;
    ConflictPromptItemFacts destinationItem;
    uint8_t dateRelation = 0u; // ConflictDateRelation
    uint8_t sizeRelation = 0u; // ConflictSizeRelation
};

struct ConflictObservedComparison; // FolderWindow.FileOperationsInternal.h

// I62 B3.3: an Issues row's lines for the two items a collision decision showed, in the decision's words.
struct ConflictObservedTexts final
{
    std::wstring source;
    std::wstring destination;
    std::wstring summary; // `Same size and date` or `One hour apart, same size`; empty otherwise
};

// The one leaf both roles name, used as a heading; empty unless a name collision names the same admitted leaf.
[[nodiscard]] std::wstring_view CommonConflictPromptName(const ConflictPromptTextFacts& facts) noexcept;
// The question the prompt asks.
[[nodiscard]] std::wstring BuildConflictPromptQuestion(const ConflictPromptTextFacts& facts);
// The item and byte counts a consent names. On a name collision, the one line said for both items
// when their size and date are the same or one hour apart. Empty when there is nothing to say.
[[nodiscard]] std::wstring BuildConflictPromptFactText(const ConflictPromptTextFacts& facts);
// One role's line on a name collision: its size (or the folder type), its date, and how it compares
// with the other item (`Newer`, `Smaller`...). The relation words are left out when the fact text
// above already speaks for both items. Empty when nothing is known about the item.
[[nodiscard]] std::wstring BuildConflictPromptItemText(const ConflictPromptTextFacts& facts, bool destinationRole);
// I62 B3.3: the role lines and the summary line of a retained comparison, built by BuildConflictPromptItemText and
// BuildConflictPromptFactText over the same facts; the Issues pane's detail shows them.
[[nodiscard]] ConflictObservedTexts BuildConflictObservedTexts(FileSystemOperation operation, const ConflictObservedComparison& observed);
// A file time as the user's short date and time; empty for an unknown time.
[[nodiscard]] std::wstring FormatConflictPromptFileTime(__int64 fileTime) noexcept;
// What a screen reader is told once when the decision is published: the waiting state, the shared name
// and the question (R8.12: the card and the decision mirror announce the same text, never both).
[[nodiscard]] std::wstring BuildConflictPromptAnnouncement(const ConflictPromptTextFacts& facts);
// The operation's card title (Copy, Move, Delete, Rename).
[[nodiscard]] UINT ConflictPromptTitleStringId(FileSystemOperation operation) noexcept;
} // namespace FileOperations
