#pragma once

#include "ShortcutDefaults.h"

namespace SettingsSave
{
[[nodiscard]] inline Common::Settings::Settings PrepareForSave(const Common::Settings::Settings& settings)
{
    Common::Settings::Settings result = settings;

    if (result.shortcuts.has_value() && ShortcutDefaults::AreShortcutsDefault(result.shortcuts.value()))
    {
        result.shortcuts.reset();
    }

    if (result.monitor.has_value())
    {
        const Common::Settings::MonitorSettings defaults{};
        const auto& monitor = result.monitor.value();

        if (monitor.menu.toolbarVisible == defaults.menu.toolbarVisible && monitor.menu.lineNumbersVisible == defaults.menu.lineNumbersVisible &&
            monitor.menu.alwaysOnTop == defaults.menu.alwaysOnTop && monitor.menu.showIds == defaults.menu.showIds &&
            monitor.menu.autoScroll == defaults.menu.autoScroll && (monitor.filter.mask & 63u) == (defaults.filter.mask & 63u) &&
            monitor.filter.preset == defaults.filter.preset)
        {
            result.monitor.reset();
        }
    }

    if (result.cache.has_value())
    {
        const auto& directoryInfo     = result.cache->directoryInfo;
        const bool wroteDirectoryInfo = (directoryInfo.maxBytes.has_value() && directoryInfo.maxBytes.value() > 0) || directoryInfo.maxWatchers.has_value() ||
                                        directoryInfo.mruWatched.has_value();
        if (! wroteDirectoryInfo)
        {
            result.cache.reset();
        }
    }

    if (result.fileOperations.has_value())
    {
        if (! Common::Settings::HasNonDefaultFileOperationsSettings(result.fileOperations.value()))
        {
            result.fileOperations.reset();
        }
    }

    if (result.compareDirectories.has_value())
    {
        const Common::Settings::CompareDirectoriesSettings defaults{};
        const auto& compare      = result.compareDirectories.value();
        const bool hasNonDefault = compare.compareSize != defaults.compareSize || compare.compareDateTime != defaults.compareDateTime ||
                                   compare.compareAttributes != defaults.compareAttributes || compare.compareContent != defaults.compareContent ||
                                   compare.compareSubdirectories != defaults.compareSubdirectories ||
                                   compare.compareSubdirectoryAttributes != defaults.compareSubdirectoryAttributes ||
                                   compare.selectSubdirsOnlyInOnePane != defaults.selectSubdirsOnlyInOnePane || compare.ignoreFiles != defaults.ignoreFiles ||
                                   compare.ignoreDirectories != defaults.ignoreDirectories || compare.keepIdenticalItems != defaults.keepIdenticalItems ||
                                   compare.showIdenticalItems != defaults.showIdenticalItems ||
                                   compare.contentCompareWorkerCount != defaults.contentCompareWorkerCount || ! compare.ignoreFilesPatterns.empty() ||
                                   ! compare.ignoreDirectoriesPatterns.empty();
        if (! hasNonDefault)
        {
            result.compareDirectories.reset();
        }
    }

    if (result.hotPaths.has_value())
    {
        const auto& hp = result.hotPaths.value();

        bool hasAnyPath = false;
        for (const auto& slot : hp.slots)
        {
            if (slot.has_value() && ! slot.value().path.empty())
            {
                hasAnyPath = true;
                break;
            }
        }

        if (! hasAnyPath && ! hp.openPrefsOnAssign)
        {
            result.hotPaths.reset();
        }
    }

    if (result.selectionMasks.has_value())
    {
        const auto& masks = result.selectionMasks.value();
        if (masks.selectHistory.empty() && masks.unselectHistory.empty() && masks.filterHistory.empty())
        {
            result.selectionMasks.reset();
        }
    }

    if (result.search.has_value())
    {
        const Common::Settings::SearchDialogSettings defaults{};
        const auto& search       = result.search.value();
        const bool hasNonDefault = ! search.recentRoots.empty() || ! search.recentNamePatterns.empty() || ! search.recentContentPatterns.empty() ||
                                   ! search.lastRoot.empty() || ! search.lastNamePattern.empty() || ! search.lastContentPattern.empty() ||
                                   search.recursive != defaults.recursive || search.includeFiles != defaults.includeFiles ||
                                   search.includeDirectories != defaults.includeDirectories || search.followSymlinks != defaults.followSymlinks ||
                                   search.matchCaseName != defaults.matchCaseName || search.matchCaseContent != defaults.matchCaseContent ||
                                   search.preferIndex != defaults.preferIndex || search.wantSnippets != defaults.wantSnippets ||
                                   search.nameMode != defaults.nameMode || search.contentMode != defaults.contentMode ||
                                   search.maxResults != defaults.maxResults || ! search.resultsGridLayout.empty();
        if (! hasNonDefault)
        {
            result.search.reset();
        }
    }

    if (result.batchRename.has_value())
    {
        const Common::Settings::BatchRenameSettings defaults{};
        const auto& batchRename  = result.batchRename.value();
        const bool hasNonDefault = ! batchRename.lastRoot.empty() || ! batchRename.recentMasks.empty() || ! batchRename.recentNameTemplates.empty() ||
                                   ! batchRename.recentSearchPatterns.empty() || ! batchRename.recentReplacePatterns.empty() ||
                                   batchRename.includeSubdirectories != defaults.includeSubdirectories || batchRename.includeFiles != defaults.includeFiles ||
                                   batchRename.includeFolders != defaults.includeFolders || batchRename.regexEnabled != defaults.regexEnabled ||
                                   batchRename.caseSensitive != defaults.caseSensitive || batchRename.wholeWords != defaults.wholeWords ||
                                   batchRename.replaceOnce != defaults.replaceOnce || batchRename.excludeExtension != defaults.excludeExtension ||
                                   batchRename.flattenSeparator != defaults.flattenSeparator || batchRename.fileNameCaseStyle != defaults.fileNameCaseStyle ||
                                   batchRename.extensionCaseStyle != defaults.extensionCaseStyle || ! batchRename.previewSortColumnId.empty() ||
                                   batchRename.previewSortDescending != defaults.previewSortDescending || ! batchRename.previewGridLayout.empty();
        if (! hasNonDefault)
        {
            result.batchRename.reset();
        }
    }

    if (result.makeFileList.has_value())
    {
        const Common::Settings::MakeFileListSettings defaults{};
        if (result.makeFileList.value() == defaults)
        {
            result.makeFileList.reset();
        }
    }

    if (result.ui.has_value())
    {
        const Common::Settings::UiSettings defaults{};
        if (result.ui.value() == defaults)
        {
            result.ui.reset();
        }
    }

    return result;
}
} // namespace SettingsSave
