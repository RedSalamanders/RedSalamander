#include "FileOperationConfirmation.h"

#include "Helpers.h"
#include "HostServices.h"
#include "resource.h"

bool ConfirmNonRevertableFileOperation(HWND owner,
                                       FileSystemOperation operation,
                                       const std::vector<std::filesystem::path>& sourcePaths,
                                       const std::filesystem::path& destinationFolder,
                                       const NonRevertableFileOperationPromptCounts& counts,
                                       HostFileOperationPromptOptions* options,
                                       std::wstring_view confirmationMessage) noexcept
{
    if (operation != FILESYSTEM_COPY && operation != FILESYSTEM_MOVE)
    {
        return true;
    }

    if (sourcePaths.empty())
    {
        return true;
    }

    auto suffixFor = [](unsigned long long count) noexcept -> std::wstring_view
    { return count == 1ull ? std::wstring_view(L"") : std::wstring_view(L"s"); };

    const unsigned long long itemCount = static_cast<unsigned long long>(sourcePaths.size());
    std::wstring what;
    if (counts.unknownCount > 0)
    {
        const std::wstring_view itemSuffix = suffixFor(itemCount);
        what                               = FormatStringResource(nullptr, IDS_FMT_FILEOPS_COUNT_ITEM, itemCount, itemSuffix);
    }
    else if (counts.fileCount > 0 && counts.folderCount > 0)
    {
        const std::wstring_view fileSuffix   = suffixFor(counts.fileCount);
        const std::wstring_view folderSuffix = suffixFor(counts.folderCount);
        what = FormatStringResource(nullptr, IDS_FMT_FILEOPS_COUNT_FILES_FOLDERS, counts.fileCount, fileSuffix, counts.folderCount, folderSuffix);
    }
    else if (counts.fileCount > 0)
    {
        const std::wstring_view fileSuffix = suffixFor(counts.fileCount);
        what                               = FormatStringResource(nullptr, IDS_FMT_FILEOPS_COUNT_FILE, counts.fileCount, fileSuffix);
    }
    else
    {
        const std::wstring_view folderSuffix = suffixFor(counts.folderCount);
        what                                 = FormatStringResource(nullptr, IDS_FMT_FILEOPS_COUNT_FOLDER, counts.folderCount, folderSuffix);
    }

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

    std::wstring fromText;
    if (sourcePaths.size() == 1u)
    {
        fromText = sourcePaths.front().wstring();
        if (counts.unknownCount == 0 && counts.folderCount == 1ull && counts.fileCount == 0ull)
        {
            fromText = ensureTrailingSeparator(std::move(fromText));
        }
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

        if (multipleParents)
        {
            fromText = LoadStringResource(nullptr, IDS_FILEOPS_LOCATION_MULTIPLE);
        }
        else if (counts.unknownCount == 0 && counts.fileCount > 0 && counts.folderCount > 0 && counts.hasSampleFile)
        {
            fromText = counts.sampleFile.wstring();
        }
        else
        {
            fromText = ensureTrailingSeparator(commonParent.wstring());
        }
    }

    std::wstring toText = ensureTrailingSeparator(destinationFolder.wstring());
    normalizeSlashes(fromText);
    normalizeSlashes(toText);

    const UINT messageId = operation == FILESYSTEM_COPY ? static_cast<UINT>(IDS_FMT_FILEOPS_CONFIRM_COPY) : static_cast<UINT>(IDS_FMT_FILEOPS_CONFIRM_MOVE);
    const std::wstring message =
        confirmationMessage.empty() ? FormatStringResource(nullptr, messageId, what, fromText, toText) : std::wstring(confirmationMessage);

    const bool isCopy = operation == FILESYSTEM_COPY;
    const UINT operationLabelId = isCopy ? static_cast<UINT>(IDS_FILEOP_OPERATION_COPY) : static_cast<UINT>(IDS_FILEOP_OPERATION_MOVE);
    const std::wstring caption = LoadStringResource(nullptr, operationLabelId);
    const HostPromptPresentation presentation = isCopy ? HOST_PROMPT_PRESENTATION_COPY : HOST_PROMPT_PRESENTATION_MOVE;
    HostPromptRequest prompt{};
    prompt.sizeBytes     = sizeof(prompt);
    prompt.scope         = (owner && IsWindow(owner)) ? HOST_ALERT_SCOPE_WINDOW : HOST_ALERT_SCOPE_APPLICATION;
    prompt.severity      = HOST_ALERT_INFO;
    prompt.buttons       = HOST_PROMPT_BUTTONS_OK_CANCEL;
    prompt.targetWindow  = (prompt.scope == HOST_ALERT_SCOPE_WINDOW) ? owner : nullptr;
    prompt.title         = caption.c_str();
    prompt.message       = message.c_str();
    prompt.defaultResult = HOST_PROMPT_RESULT_OK;
    prompt.presentation  = presentation;
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
    NonRevertableFileOperationPromptCounts counts{};
    counts.unknownCount = static_cast<unsigned long long>(sourcePaths.size());
    return ConfirmNonRevertableFileOperation(owner, operation, sourcePaths, destinationFolder, counts);
}
