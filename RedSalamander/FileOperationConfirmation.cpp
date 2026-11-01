#include "FileOperationConfirmation.h"

#include "Helpers.h"
#include "HostServices.h"
#include "resource.h"

std::wstring FormatFileOperationCountNoun(const FileOperationCountNoun noun, const unsigned long long count) noexcept
{
    const Localization::CardinalPluralCategory category = Localization::SelectCardinalPluralCategory(Localization::GetEffectiveResourceCulture(nullptr), count);
    UINT resourceId                                     = 0u;
    switch (noun)
    {
        case FileOperationCountNoun::Item:
            resourceId = static_cast<UINT>(category == Localization::CardinalPluralCategory::One   ? IDS_FMT_FILEOPS_COUNT_ITEM_ONE
                                           : category == Localization::CardinalPluralCategory::Few ? IDS_FMT_FILEOPS_COUNT_ITEM_FEW
                                                                                                   : IDS_FMT_FILEOPS_COUNT_ITEM);
            break;
        case FileOperationCountNoun::File:
            resourceId = static_cast<UINT>(category == Localization::CardinalPluralCategory::One   ? IDS_FMT_FILEOPS_COUNT_FILE_ONE
                                           : category == Localization::CardinalPluralCategory::Few ? IDS_FMT_FILEOPS_COUNT_FILE_FEW
                                                                                                   : IDS_FMT_FILEOPS_COUNT_FILE);
            break;
        case FileOperationCountNoun::Folder:
            resourceId = static_cast<UINT>(category == Localization::CardinalPluralCategory::One   ? IDS_FMT_FILEOPS_COUNT_FOLDER_ONE
                                           : category == Localization::CardinalPluralCategory::Few ? IDS_FMT_FILEOPS_COUNT_FOLDER_FEW
                                                                                                   : IDS_FMT_FILEOPS_COUNT_FOLDER);
            break;
    }

    return resourceId != 0u ? FormatStringResource(nullptr, resourceId, count) : std::wstring{};
}

std::wstring FormatFileOperationConfirmationSource(const std::vector<std::filesystem::path>& sourcePaths,
                                                   const NonRevertableFileOperationPromptCounts& counts,
                                                   const FileSystemPathIdentity* sourceIdentity) noexcept
{
    if (sourcePaths.empty())
        return {};

    if (sourcePaths.size() == 1u)
    {
        const std::wstring& source = sourcePaths.front().native();
        return sourceIdentity != nullptr && counts.unknownCount == 0u && counts.folderCount == 1u && counts.fileCount == 0u
                   ? JoinFileSystemPath(*sourceIdentity, source, {})
                   : source;
    }

    std::wstring commonParent;
    bool canGroup = sourceIdentity != nullptr && sourceIdentity->pathTextStableIdentity &&
                    TryGetFileSystemParentPath(*sourceIdentity, sourcePaths.front().native(), commonParent);
    for (size_t index = 1u; canGroup && index < sourcePaths.size(); ++index)
    {
        std::wstring parent;
        canGroup = TryGetFileSystemParentPath(*sourceIdentity, sourcePaths[index].native(), parent) && EquivalentPath(*sourceIdentity, commonParent, parent);
    }
    if (canGroup)
    {
        if (counts.unknownCount == 0u && counts.fileCount > 0u && counts.folderCount > 0u && counts.hasSampleFile)
            return counts.sampleFile.native();
        return JoinFileSystemPath(*sourceIdentity, commonParent, {});
    }

    // Do not invent a parent for opaque keys or case/provider-distinct locations.
    // Keep the existing bounded summary for large selections.
    return LoadStringResource(nullptr, IDS_FILEOPS_LOCATION_MULTIPLE);
}

bool ConfirmNonRevertableFileOperation(HWND owner,
                                       FileSystemOperation operation,
                                       const std::vector<std::filesystem::path>& sourcePaths,
                                       const std::filesystem::path& destinationFolder,
                                       const NonRevertableFileOperationPromptCounts& counts,
                                       HostFileOperationPromptOptions* options,
                                       std::wstring_view confirmationMessage,
                                       const FileSystemPathIdentity* sourceIdentity,
                                       const FileSystemPathIdentity* destinationIdentity,
                                       std::wstring_view appendedNotice) noexcept
{
    if (operation != FILESYSTEM_COPY && operation != FILESYSTEM_MOVE)
    {
        return true;
    }

    if (sourcePaths.empty())
    {
        return true;
    }

    const unsigned long long itemCount = static_cast<unsigned long long>(sourcePaths.size());
    std::wstring what;
    if (counts.unknownCount > 0)
    {
        what = FormatFileOperationCountNoun(FileOperationCountNoun::Item, itemCount);
    }
    else if (counts.fileCount > 0 && counts.folderCount > 0)
    {
        what = FormatStringResource(nullptr,
                                    IDS_FMT_FILEOPS_COUNT_FILES_FOLDERS,
                                    FormatFileOperationCountNoun(FileOperationCountNoun::File, counts.fileCount),
                                    FormatFileOperationCountNoun(FileOperationCountNoun::Folder, counts.folderCount));
    }
    else if (counts.fileCount > 0)
    {
        what = FormatFileOperationCountNoun(FileOperationCountNoun::File, counts.fileCount);
    }
    else
    {
        what = FormatFileOperationCountNoun(FileOperationCountNoun::Folder, counts.folderCount);
    }

    const std::wstring fromText = FormatFileOperationConfirmationSource(sourcePaths, counts, sourceIdentity);
    const std::wstring toText =
        destinationIdentity != nullptr ? JoinFileSystemPath(*destinationIdentity, destinationFolder.native(), {}) : destinationFolder.native();

    const UINT messageId = operation == FILESYSTEM_COPY ? static_cast<UINT>(IDS_FMT_FILEOPS_CONFIRM_COPY) : static_cast<UINT>(IDS_FMT_FILEOPS_CONFIRM_MOVE);
    std::wstring message = confirmationMessage.empty() ? FormatStringResource(nullptr, messageId, what, fromText, toText) : std::wstring(confirmationMessage);
    if (! appendedNotice.empty())
    {
        message.append(L"\r\n\r\n");
        message.append(appendedNotice);
    }

    const bool isCopy                         = operation == FILESYSTEM_COPY;
    const UINT operationLabelId               = isCopy ? static_cast<UINT>(IDS_FILEOP_OPERATION_COPY) : static_cast<UINT>(IDS_FILEOP_OPERATION_MOVE);
    const std::wstring caption                = LoadStringResource(nullptr, operationLabelId);
    const HostPromptPresentation presentation = isCopy ? HOST_PROMPT_PRESENTATION_COPY : HOST_PROMPT_PRESENTATION_MOVE;
    HostPromptRequest prompt{};
    prompt.sizeBytes            = sizeof(prompt);
    prompt.scope                = (owner && IsWindow(owner)) ? HOST_ALERT_SCOPE_WINDOW : HOST_ALERT_SCOPE_APPLICATION;
    prompt.severity             = HOST_ALERT_INFO;
    prompt.buttons              = HOST_PROMPT_BUTTONS_OK_CANCEL;
    prompt.targetWindow         = (prompt.scope == HOST_ALERT_SCOPE_WINDOW) ? owner : nullptr;
    prompt.title                = caption.c_str();
    prompt.message              = message.c_str();
    prompt.defaultResult        = HOST_PROMPT_RESULT_OK;
    prompt.presentation         = presentation;
    prompt.fileOperationOptions = options;

    HostPromptResult promptResult = HOST_PROMPT_RESULT_NONE;
    const HRESULT hr              = HostShowPrompt(prompt, nullptr, &promptResult);
    if (FAILED(hr))
    {
        return false;
    }

    return promptResult == HOST_PROMPT_RESULT_OK;
}

bool ConfirmNonRevertableFileOperation(HWND owner,
                                       IFileSystem* /*fileSystem*/,
                                       FileSystemOperation operation,
                                       const std::vector<std::filesystem::path>& sourcePaths,
                                       const std::filesystem::path& destinationFolder) noexcept
{
    // This legacy direct-operation test fallback has no captured endpoint
    // contracts. The adapter alone cannot identify clipboard/drop sources;
    // preserve raw display rather than infer Windows rules or query new facts.
    NonRevertableFileOperationPromptCounts counts{};
    counts.unknownCount = static_cast<unsigned long long>(sourcePaths.size());
    return ConfirmNonRevertableFileOperation(owner, operation, sourcePaths, destinationFolder, counts);
}
