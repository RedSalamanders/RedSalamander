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
            monitor.menu.autoScroll == defaults.menu.autoScroll && (monitor.filter.mask & 31u) == (defaults.filter.mask & 31u) &&
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
        const Common::Settings::FileOperationsSettings defaults{};
        const auto& fileOperations = result.fileOperations.value();
        const bool hasNonDefault =
            fileOperations.autoDismissSuccess != defaults.autoDismissSuccess || fileOperations.maxDiagnosticsLogFiles != defaults.maxDiagnosticsLogFiles ||
            fileOperations.diagnosticsInfoEnabled != defaults.diagnosticsInfoEnabled ||
            fileOperations.diagnosticsDebugEnabled != defaults.diagnosticsDebugEnabled || fileOperations.maxIssueReportFiles.has_value() ||
            fileOperations.maxDiagnosticsInMemory.has_value() || fileOperations.maxDiagnosticsPerFlush.has_value() ||
            fileOperations.diagnosticsFlushIntervalMs.has_value() || fileOperations.diagnosticsCleanupIntervalMs.has_value();
        if (! hasNonDefault)
        {
            result.fileOperations.reset();
        }
    }

    if (result.compareDirectories.has_value())
    {
        const Common::Settings::CompareDirectoriesSettings defaults{};
        const auto& compare = result.compareDirectories.value();
        const bool hasNonDefault = compare.compareSize != defaults.compareSize || compare.compareDateTime != defaults.compareDateTime ||
                                   compare.compareAttributes != defaults.compareAttributes || compare.compareContent != defaults.compareContent ||
                                   compare.compareSubdirectories != defaults.compareSubdirectories ||
                                   compare.compareSubdirectoryAttributes != defaults.compareSubdirectoryAttributes ||
                                   compare.selectSubdirsOnlyInOnePane != defaults.selectSubdirsOnlyInOnePane || compare.ignoreFiles != defaults.ignoreFiles ||
                                   compare.ignoreDirectories != defaults.ignoreDirectories || compare.showIdenticalItems != defaults.showIdenticalItems ||
                                   ! compare.ignoreFilesPatterns.empty() || ! compare.ignoreDirectoriesPatterns.empty();
        if (! hasNonDefault)
        {
            result.compareDirectories.reset();
        }
    }

    return result;
}
} // namespace SettingsSave
