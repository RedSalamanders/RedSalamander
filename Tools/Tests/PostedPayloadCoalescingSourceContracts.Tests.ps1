Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)

function Get-RSText {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    return Get-Content -LiteralPath (Join-Path $repoRoot $Path) -Raw
}

Describe 'Posted payload coalescing source contracts' {
    It 'checks the unfiltered queue head before removing an exact operation-key match' {
        $source = Get-RSText -Path 'Common\Helpers.h'

        $source | Should Match 'TakeAndCoalesceContiguousPostedPayloads'
        $source | Should Match 'PeekMessageW\(&queuedMessage,\s*nullptr,\s*0,\s*0,\s*PM_NOREMOVE\)'
        $source | Should Match 'queuedMessage\.hwnd\s*!=\s*hwnd'
        $source | Should Match 'queuedMessage\.message\s*!=\s*message'
        $source | Should Match 'queuedMessage\.wParam\s*!=\s*operationKey'
        $source | Should Match 'TakeMessagePayload<T>\(queuedMessage\.lParam\)'
    }

    It 'keeps Compare progress drains on the shared keyed helper' {
        $source = Get-RSText -Path 'RedSalamander\CompareDirectoriesWindow.Progress.cpp'

        ([regex]::Matches($source, 'TakeAndCoalesceContiguousPostedPayloads<').Count) | Should Be 2
        $source | Should Not Match 'PeekMessageW\('
        $source | Should Match 'kCompareDirectoriesScanProgress,\s*operationKey,\s*std::move\(payload\)'
        $source | Should Match 'kCompareDirectoriesContentProgress,\s*operationKey,\s*std::move\(payload\)'
    }

    It 'keeps Find result and progress drains keyed while preserving their budgets' {
        $source = Get-RSText -Path 'RedSalamander\FindFilesWindow.cpp'

        ([regex]::Matches($source, 'TakeAndCoalesceContiguousPostedPayloads<').Count) | Should Be 2
        $source | Should Match 'kFindSearchResults,\s*operationKey,\s*std::move\(payload\)'
        $source | Should Match 'kFindSearchProgress,\s*operationKey,\s*std::move\(payload\)'
        $source | Should Match 'kFindSearchComplete,\s*operationKey,\s*std::move\(complete\)'
        $source | Should Match 'drainedPayloadCount\s*<\s*kResultsDrainMaxMessages'
        $source | Should Match 'current\.results\.size\(\)\s*<\s*kResultsDrainMaxRecords'
    }
}
