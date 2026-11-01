<#
.SYNOPSIS
    Audits remaining Win32 UI dependencies and their reviewed residual allowlist.

.DESCRIPTION
    Scans owned source roots for native font, GDI text, device-context, and visible
    control patterns. A residual is allowed only by an exact category, source path, and
    line-pattern entry with visibility, owner, reason, and exit condition. Use
    -FailOnFindings to make unclassified findings fail automation.

.PARAMETER RepoRoot
    Repository root to scan. Defaults to the parent of the Tools directory.

.PARAMETER AsMarkdown
    Emits a grouped Markdown summary and finding tables.

.PARAMETER FailOnFindings
    Throws when one or more findings do not match the reviewed residual allowlist.

.OUTPUTS
    Structured finding objects by default, with the category summary on the information
    stream. Markdown text with AsMarkdown.

.NOTES
    Prerequisites: repository product/test source. Side effects: none; scanning is read-only. FailOnFindings produces a nonzero exit for unreviewed dependencies. Primary use is Win32 UI migration closeout and source-policy review.

.EXAMPLE
    .\Tools\Audit-RemainingWin32UiDependencies.ps1 -FailOnFindings

.EXAMPLE
    .\Tools\Audit-RemainingWin32UiDependencies.ps1 -AsMarkdown
#>
[CmdletBinding()]
param(
    [string]$RepoRoot = (Resolve-Path (Join-Path $PSScriptRoot '..')).Path,
    [switch]$AsMarkdown,
    [switch]$FailOnFindings
)

$ErrorActionPreference = 'Stop'
Import-Module (Join-Path $PSScriptRoot 'Modules\Auditing\RepositorySourceScanner.psm1') -Force

$sourceRoots = @(
    'Common',
    'RedSalamander',
    'Plugins',
    'Tests'
)

$fileExtensions = @('*.c', '*.cpp', '*.h', '*.hpp', '*.inl', '*.rc')

$excludedPathFragments = @(
    '\.build\',
    '\Specs\TestRuns\',
    '\vcpkg_installed\',
    '\.git\'
)

$patterns = @(
    @{ Category = 'HFONT handle'; Pattern = '\bHFONT\b' },
    @{ Category = 'LOGFONT bridge'; Pattern = '\bLOGFONTW?\b' },
    @{ Category = 'GDI font creation'; Pattern = '\bCreateFontIndirectW?\s*\(|\bCreateFontW\s*\(' },
    @{ Category = 'Native font message'; Pattern = 'WM_SETFONT|WM_GETFONT' },
    @{ Category = 'GDI text draw'; Pattern = '(?<!->)\bDrawTextW\s*\(|\bTextOutW\s*\(|\bExtTextOutW\s*\(' },
    @{ Category = 'HDC text/selection bridge'; Pattern = 'GetDC\(|SelectObject\(' },
    @{ Category = 'Native status/common control'; Pattern = 'STATUSCLASSNAMEW?|WC_LISTVIEWW?|WC_TREEVIEWW?|WC_HEADERW?|WC_TABCONTROLW?|TOOLBARCLASSNAMEW?|TRACKBAR_CLASSW?|PROGRESS_CLASSW?' },
    @{ Category = 'Native visible control creation'; Pattern = 'CreateWindowExW\([^`r`n]*(STATIC|BUTTON|EDIT|COMBOBOX|LISTBOX|SysListView32|SysTreeView32)' }
)

$allowedResidualDependencies = @(
    @{
        Category      = 'HDC text/selection bridge'
        Path          = 'RedSalamander\FolderView.Icons.cpp'
        LinePattern   = 'GetDC\(nullptr\)'
        Visibility    = 'test-only bitmap interop'
        Owner         = 'FolderView thumbnail self-tests'
        Reason        = 'ENABLE_TESTS synthetic thumbnail generation uses a screen DC only to create deterministic DIB thumbnails; it does not render native text, fonts, or controls.'
        ExitCondition = 'Replace if thumbnail self-tests gain a non-HDC bitmap factory helper.'
    },
    @{
        Category      = 'HDC text/selection bridge'
        Path          = 'RedSalamander\FolderView.Rendering.cpp'
        LinePattern   = 'GetDC\(nullptr\)|SelectObject\(hdcMem\.get\(\), hBitmap\.get\(\)\)'
        Visibility    = 'shell icon bitmap interop'
        Owner         = 'FolderView icon renderer'
        Reason        = 'Shortcut overlay extraction converts a shell HICON into a D2D bitmap; it is not a native text/font/control path.'
        ExitCondition = 'Replace when shell-icon conversion is performed through a non-HDC imaging path.'
    },
    @{
        Category      = 'HDC text/selection bridge'
        Path          = 'RedSalamander\IconCache.cpp'
        LinePattern   = 'GetDC\(nullptr\)|SelectObject\(memoryDc\.get\(\), (bitmap|dib)\.get\(\)\)'
        Visibility    = 'shell icon bitmap interop'
        Owner         = 'IconCache'
        Reason        = 'IconCache uses HDC/DIB interop to normalize shell icons before Direct2D rendering; it does not draw app-owned visible text.'
        ExitCondition = 'Replace when icon extraction/conversion has a WIC/Direct2D-only implementation.'
    },
    @{
        Category      = 'HDC text/selection bridge'
        Path          = 'RedSalamander\NavigationViewInternal.h'
        LinePattern   = 'SelectObject\((destMem|srcMem)\.get\(\), (destDib|srcDib)\.get\(\)\)'
        Visibility    = 'bitmap alpha-blend compatibility'
        Owner         = 'NavigationView image renderer'
        Reason        = 'The compatibility AlphaBlend fallback copies pixels through DIB sections; it does not render native text, fonts, or controls.'
        ExitCondition = 'Replace when the fallback alpha-blend path moves to Direct2D/WIC or is no longer needed.'
    },
    @{
        Category      = 'HDC text/selection bridge'
        Path          = 'RedSalamander\SelfTest\Commands\Commands.Tests.ViewCommands.cpp'
        LinePattern   = 'GetDC\(hwnd\)'
        Visibility    = 'test-only'
        Owner         = 'Commands self-tests'
        Reason        = 'Pixel probes validate rendered DxUi surfaces in deterministic command self-tests.'
        ExitCondition = 'Replace with a non-HDC render capture helper if one becomes available.'
    },
    @{
        Category      = 'GDI font creation'
        Path          = 'Tests\TestSupport\TestSupport.Tests.DirectedSelfTestInputWarning.h'
        LinePattern   = 'CreateFontW\('
        Visibility    = 'test-only operator warning'
        Owner         = 'Shared foreground-input selftests'
        Reason        = 'The shared warning window is shown only while a directed desktop-input selftest is active; it is not product UI.'
        ExitCondition = 'Replace if the directed-input harness gains a DirectWrite warning surface outside the product process.'
    },
    @{
        Category      = 'Native font message'
        Path          = 'Tests\TestSupport\TestSupport.Tests.DirectedSelfTestInputWarning.h'
        LinePattern   = 'WM_SETFONT'
        Visibility    = 'test-only operator warning'
        Owner         = 'Shared foreground-input selftests'
        Reason        = 'The shared warning window is shown only while a directed desktop-input selftest is active; it is not product UI.'
        ExitCondition = 'Replace if the directed-input harness gains a DirectWrite warning surface outside the product process.'
    },
    @{
        Category      = 'HDC text/selection bridge'
        Path          = 'RedSalamander\Ui\AlertOverlayWindow.cpp'
        LinePattern   = 'GetDC\(nullptr\)|SelectObject\(memoryDc\.get\(\), bitmap\.get\(\)\)'
        Visibility    = 'visual bitmap interop'
        Owner         = 'AlertOverlayWindow'
        Reason        = 'Modal alert overlays use a memory DC only to capture the owner backdrop before Direct2D scrim composition; it does not render native text, fonts, or controls.'
        ExitCondition = 'Replace when the alert-overlay backdrop path moves to DirectComposition/WIC-only capture or the captured-backdrop effect is removed.'
    },
    @{
        Category      = 'Native visible control creation'
        Path          = 'Plugins\ViewerVLC\ViewerVLC.cpp'
        LinePattern   = 'WS_CHILD \| WS_VISIBLE \| SS_BLACKRECT, 0, 0, 24, 24'
        Visibility    = 'test-only synthetic child'
        Owner         = 'ViewerVLC wheel-forwarding self-tests'
        Reason        = 'ENABLE_TESTS creates a tiny synthetic Static child only to exercise VLC child-window wheel forwarding; it is not production app chrome.'
        ExitCondition = 'Remove if the wheel-forwarding test can use a non-HWND child probe.'
    },
    @{
        Category      = 'Native visible control creation'
        Path          = 'RedSalamander\SelfTest\Commands\Commands.Tests.Connections.cpp'
        LinePattern   = 'host-service-stale-target'
        Visibility    = 'test-only destroyed target'
        Owner         = 'Commands host-services self-tests'
        Reason        = 'The non-visible one-pixel Static HWND is destroyed immediately so the host-services test can prove safe fallback from a stale owner window.'
        ExitCondition = 'Remove if the stale-owner fallback can be exercised with a non-HWND lifetime probe.'
    },
    @{
        Category      = 'Native visible control creation'
        Path          = 'RedSalamander\SelfTest\Commands\Commands.Tests.Search.cpp'
        LinePattern   = 'CreateWindowExW\(0, L"STATIC", L"", WS_(POPUP|CHILD)'
        Visibility    = 'test-only hidden queue probe'
        Owner         = 'Find dialog command self-tests'
        Reason        = 'The queue-order test creates hidden Static probe windows without WS_VISIBLE to verify parent/child input-drain ordering.'
        ExitCondition = 'Remove if the drain-order test gains a non-HWND message queue probe.'
    },
    @{
        Category      = 'HFONT handle'
        Path          = 'Tests\ViewerPETests\ViewerPE.Tests.Runner.cpp'
        LinePattern   = 'legacyVisibleHfontSurfaceCount|HFONT text surfaces'
        Visibility    = 'test-only assertion text'
        Owner         = 'ViewerPETests'
        Reason        = 'The test references HFONT only in diagnostics and a debug counter name while asserting that visible legacy HFONT text surfaces are absent.'
        ExitCondition = 'Rename the diagnostic counter if the audit eventually excludes identifiers and string literals.'
    },
    @{
        Category      = 'GDI font creation'
        Path          = 'RedSalamander\FolderWindow.FileOperations.Popup.cpp'
        LinePattern   = 'CreateFontIndirectW\(&metrics\.lfMessageFont\)'
        Visibility    = 'production emergency failure surface'
        Owner         = 'File Operations popup failure surface'
        Reason        = 'Only `EnterFailureSurface` uses this font after the DxUi host is absent or repeated frame-preparation failures cross the declared threshold; it keeps the error text and Cancel All/Close recovery actions readable during that degraded state.'
        ExitCondition = 'Replace when the failure surface renders its diagnostic text and recovery controls through a supported DxUi/DirectWrite surface without native font fallback.'
    },
    @{
        Category      = 'HFONT handle'
        Path          = 'RedSalamander\FolderWindow.FileOperations.Popup.cpp'
        LinePattern   = 'const HFONT font = _failureFont \? _failureFont\.get\(\) : static_cast<HFONT>\(GetStockObject\(DEFAULT_GUI_FONT\)\);'
        Visibility    = 'production emergency failure surface'
        Owner         = 'File Operations popup failure surface'
        Reason        = 'Only the degraded failure surface selects a DPI-matched message font and, if creation fails, the stock GUI font as last-resort support for its native diagnostic controls.'
        ExitCondition = 'Replace when the failure surface renders its diagnostic text and recovery controls through a supported DxUi/DirectWrite surface without native font fallback.'
    },
    @{
        Category      = 'Native font message'
        Path          = 'RedSalamander\FolderWindow.FileOperations.Popup.cpp'
        LinePattern   = 'SendMessageW\(child, WM_SETFONT, reinterpret_cast<WPARAM>\(font\), TRUE\);'
        Visibility    = 'production emergency failure surface'
        Owner         = 'File Operations popup failure surface'
        Reason        = 'The native children receive that font only while the popup is in its detached/repeated-frame-failure recovery surface, keeping diagnostic text and recovery labels usable when the hosted controls cannot be used.'
        ExitCondition = 'Replace when the failure surface renders its diagnostic text and recovery controls through a supported DxUi/DirectWrite surface without native font messaging.'
    },
    @{
        Category      = 'HDC text/selection bridge'
        Path          = 'RedSalamander\SelfTest\Commands\Commands.Tests.FileOps.cpp'
        LinePattern   = 'wil::unique_hdc_window dc\{GetDC\(nullptr\)\};'
        Visibility    = 'test-only clipboard bitmap serialization'
        Owner         = 'Commands File Operations clipboard self-tests'
        Reason        = 'The clipboard contract fixture materializes CF_BITMAP pixels with GetDIBits into a bounded owned byte snapshot; the DC is used for bitmap serialization, not text, font, or control rendering.'
        ExitCondition = 'Replace when CF_BITMAP fixture serialization uses a bitmap-only snapshot helper with the same byte-level clipboard contract.'
    },
    @{
        Category      = 'Native visible control creation'
        Path          = 'RedSalamander\SelfTest\Commands\Commands.Tests.Dialogs.cpp'
        LinePattern   = 'drainWindow\(CreateWindowExW\(0, L"STATIC", L"", 0, 0, 0, 0, 0, HWND_MESSAGE'
        Visibility    = 'test-only message-only payload receiver'
        Owner         = 'Commands Change Case payload self-tests'
        Reason        = 'The message-only Static window owns a real posted-payload registry so the Change Case tests can verify failed-post cleanup and payload draining.'
        ExitCondition = 'Remove when the posted-payload registry has an injectable non-HWND dispatch test surface.'
    },
    @{
        Category      = 'Native visible control creation'
        Path          = 'RedSalamander\SelfTest\FileOperations\FolderWindow.FileOperations.Tests.PathDetailFocus.cpp'
        LinePattern   = 'ownershipChild\(CreateWindowExW\(0, L"STATIC", L"", WS_CHILD'
        Visibility    = 'test-only hidden focus-ownership child'
        Owner         = 'File Operations path-detail focus self-tests'
        Reason        = 'The hidden child supplies a genuine HWND ancestry relationship to the path-detail focus policy matrix; the fixture makes no SetFocus or foreground writes.'
        ExitCondition = 'Replace when the focus-ownership contract can be exercised with a non-HWND ancestry probe.'
    },
    @{
        Category      = 'Native visible control creation'
        Path          = 'RedSalamander\SelfTest\FileOperations\FolderWindow.FileOperations.Tests.Phases10_13.cpp'
        LinePattern   = 'state\.c1DeferredCloseTarget = CreateWindowExW\(0, L"STATIC", L"", 0, 0, 0, 0, 0, HWND_MESSAGE'
        Visibility    = 'test-only message-only close receiver'
        Owner         = 'File Operations C1 deferred-close self-test'
        Reason        = 'The message-only target receives and destroys on WM_CLOSE so the C1 lifecycle test can observe deferred application-close ordering without a visible window.'
        ExitCondition = 'Replace when the C1 close-order contract has a non-HWND message-target seam.'
    },
    @{
        Category      = 'Native visible control creation'
        Path          = 'RedSalamander\SelfTest\FileOperations\FolderWindow.FileOperations.Tests.QueuedAdmission.cpp'
        LinePattern   = 'tokenWindow\(CreateWindowExW\(0u, L"STATIC", L"", 0u, 0, 0, 0, 0, HWND_MESSAGE'
        Visibility    = 'test-only message-only payload registry window'
        Owner         = 'File Operations queued-admission self-tests'
        Reason        = 'The message-only window registers and retires a real posted-payload token, allowing the queued-admission case to replay a stale token through the production registry.'
        ExitCondition = 'Replace when the stale-token contract can run against an injectable payload registry without an HWND.'
    },
    @{
        Category      = 'Native visible control creation'
        Path          = 'RedSalamander\SelfTest\FileOperations\FolderWindow.FileOperations.Tests.Runner.cpp'
        LinePattern   = 'CreateWindowExW\(0, L"STATIC", L"Owned completion teardown fixture", WS_POPUP'
        Visibility    = 'test-only hidden FolderWindow owner'
        Owner         = 'RunFileOpsCompletionOwnerDestroyChecks self-test'
        Reason        = 'The hidden popup HWND is the real parent for a secondary FolderWindow in the owned-completion teardown regression; WS_POPUP is used without WS_VISIBLE.'
        ExitCondition = 'Replace when the teardown ownership contract can create its FolderWindow against a non-window host fixture.'
    },
    @{
        Category      = 'Native visible control creation'
        Path          = 'RedSalamander\SelfTest\FileOperations\FolderWindow.FileOperations.Tests.Runner.cpp'
        LinePattern   = '^target = CreateWindowExW\(0, L"STATIC", L"", 0, 0, 0, 0, 0, HWND_MESSAGE'
        Visibility    = 'test-only message-only deferred-close receiver'
        Owner         = 'RunFileOpsDeferredExitRefusesNewWorkChecks self-test'
        Reason        = 'The message-only target receives WM_CLOSE and is destroyed to witness the resumed close path while exercising deferred exit ordering without visible product UI.'
        ExitCondition = 'Replace when deferred exit-close ordering can be tested through a non-HWND close-target seam.'
    },
    @{
        Category      = 'Native visible control creation'
        Path          = 'RedSalamander\SelfTest\FileOperations\FolderWindow.FileOperations.Tests.Runner.cpp'
        LinePattern   = 'HWND closeTarget = CreateWindowExW\(0, L"STATIC", L"", 0, 0, 0, 0, 0, HWND_MESSAGE'
        Visibility    = 'test-only message-only compare-close receiver'
        Owner         = 'RunFileOpsCompareWindowCloseConsentChecks self-test'
        Reason        = 'The message-only target receives WM_CLOSE so the Compare-window close-consent case can verify the resumed application-close path without visible product UI.'
        ExitCondition = 'Replace when Compare-window close-consent ordering can be tested through a non-HWND close-target seam.'
    },
    @{
        Category      = 'Native visible control creation'
        Path          = 'RedSalamander\Ui\AlertOverlayWindow.cpp'
        LinePattern   = 'CreateWindowExW\(0, L"STATIC", L"foreign HWND reuse witness", WS_POPUP'
        Visibility    = 'test-only hidden UIA identity probe'
        Owner         = 'AlertOverlay UI Automation identity self-tests'
        Reason        = 'The hidden popup supplies a real foreign HWND so the AlertOverlay provider test can prove a stale provider rejects a reused window handle.'
        ExitCondition = 'Replace when provider identity reuse can be tested through a deterministic non-HWND handle-identity seam.'
    },
    @{
        Category      = 'Native visible control creation'
        Path          = 'RedSalamander\Ui\AlertOverlayWindow.cpp'
        LinePattern   = 'CreateWindowExW\(0, L"STATIC", L"", WS_POPUP \| WS_VISIBLE, originX, originY, 360, 240'
        Visibility    = 'test-only visible tracking anchor'
        Owner         = 'AlertOverlay tracking self-tests'
        Reason        = 'The visible synthetic anchor lets the tracking matrix exercise real native owner geometry and visible-window tracking transitions.'
        ExitCondition = 'Replace when tracking behavior can use an owned application test window with equivalent geometry and lifecycle coverage.'
    },
    @{
        Category      = 'Native visible control creation'
        Path          = 'RedSalamander\Ui\AlertOverlayWindow.cpp'
        LinePattern   = 'hiddenTarget\(CreateWindowExW\(0, L"STATIC", L"", WS_CHILD, 0, 0, 8, 8, host\.get\(\)'
        Visibility    = 'test-only hidden focus-return target'
        Owner         = 'AlertOverlay modeless focus self-tests'
        Reason        = 'The hidden child is a real HWND target used to verify focus-return fallback and ownership predicates without making the target visible.'
        ExitCondition = 'Replace when focus-return target liveness and ownership can be tested with a non-HWND window identity seam.'
    },
    @{
        Category      = 'Native visible control creation'
        Path          = 'RedSalamander\Ui\AlertOverlayWindow.cpp'
        LinePattern   = 'target\(CreateWindowExW\(0, L"STATIC", L"", WS_CHILD \| WS_VISIBLE, 0, 0, 8, 8, host\.get\(\)'
        Visibility    = 'test-only visible focus-return target'
        Owner         = 'AlertOverlay modeless focus self-tests'
        Reason        = 'The visible child is captured as the blocking alert focus-return target, proving real child-window focus restoration after dismissal.'
        ExitCondition = 'Replace when focus restoration can be validated with an owned application child window and equivalent visibility/focus observations.'
    },
    @{
        Category      = 'Native visible control creation'
        Path          = 'Tests\PluginContractTests\PluginContract.Tests.Runner.cpp'
        LinePattern   = 'wil::unique_hwnd parent\(CreateWindowExW\(0u, L"STATIC", L"", WS_OVERLAPPED'
        Visibility    = 'test-only hidden Terminal parent window'
        Owner         = 'Terminal plugin contract tests'
        Reason        = 'The hidden overlapped HWND supplies a valid parent to Terminal.Open while the isolated plugin contract checks resource-owner activation and the unsupported-location path.'
        ExitCondition = 'Remove when the Terminal.Open contract permits this diagnostic case to use a parentless context.'
    },
    @{
        Category      = 'Native visible control creation'
        Path          = 'Tests\ProductUiTests\ProductUi.Tests.Runner.cpp'
        LinePattern   = 'wil::unique_hwnd messageWindow\(CreateWindowExW\(0u, L"STATIC", L"waiting", 0u, 0, 0, 0, 0, HWND_MESSAGE'
        Visibility    = 'test-only message-only pump window'
        Owner         = 'ProductUi message-pump contract tests'
        Reason        = 'The message-only window gives the shared PumpMessagesUntil contract a real thread queue and HWND message to dispatch without creating visible UI.'
        ExitCondition = 'Replace when the message-pump contract can be validated against an injected queue implementation with the same dispatch behavior.'
    }
)

function Get-AllowlistEntry
{
    param(
        [Parameter(Mandatory = $true)]
        [pscustomobject]$Finding
    )

    foreach ($entry in $allowedResidualDependencies)
    {
        if ($entry.Category -ne $Finding.Category)
        {
            continue
        }
        if ($entry.Path -ne $Finding.Path)
        {
            continue
        }
        if ($Finding.Line -notmatch $entry.LinePattern)
        {
            continue
        }

        return $entry
    }

    return $null
}

$rawResults = @(Find-RSRepositorySourceMatches `
    -RepositoryRoot $RepoRoot `
    -SourceRoots $sourceRoots `
    -Extensions $fileExtensions `
    -Patterns $patterns `
    -ExcludedPathFragments $excludedPathFragments |
    ForEach-Object {
        [pscustomobject]@{
            Category = $_.PatternName
            Path = $_.Path
            LineNumber = $_.LineNumber
            Line = $_.Line
        }
    })

$results = foreach ($finding in @($rawResults | Sort-Object Category, Path, LineNumber))
{
    $allowlistEntry = Get-AllowlistEntry -Finding $finding
    [PSCustomObject]@{
        Category      = $finding.Category
        Path          = $finding.Path
        LineNumber    = $finding.LineNumber
        Line          = $finding.Line
        Allowed       = $null -ne $allowlistEntry
        Visibility    = if ($allowlistEntry) { $allowlistEntry.Visibility } else { '' }
        Owner         = if ($allowlistEntry) { $allowlistEntry.Owner } else { '' }
        Reason        = if ($allowlistEntry) { $allowlistEntry.Reason } else { '' }
        ExitCondition = if ($allowlistEntry) { $allowlistEntry.ExitCondition } else { '' }
    }
}

$results = @($results)
$unallowedResults = @($results | Where-Object { ! $_.Allowed })

if ($AsMarkdown)
{
    Write-Output '# Remaining Win32 UI Dependency Audit'
    Write-Output ''
    Write-Output "| Category | Count | Allowed | Unallowed |"
    Write-Output "| --- | ---: | ---: | ---: |"
    foreach ($group in ($results | Group-Object Category | Sort-Object Name))
    {
        $allowedCount = @($group.Group | Where-Object { $_.Allowed }).Count
        Write-Output "| $($group.Name) | $($group.Count) | $allowedCount | $($group.Count - $allowedCount) |"
    }

    Write-Output ''
    foreach ($group in ($results | Group-Object Category | Sort-Object Name))
    {
        Write-Output "## $($group.Name)"
        Write-Output ''
        Write-Output '| Path | Line | Allowed | Visibility | Snippet |'
        Write-Output '| --- | ---: | --- | --- | --- |'
        foreach ($item in ($group.Group | Sort-Object Path, LineNumber))
        {
            $snippet = ($item.Line -replace '\|', '\|')
            $allowed = if ($item.Allowed) { 'yes' } else { 'no' }
            Write-Output "| ``$($item.Path)`` | $($item.LineNumber) | $allowed | $($item.Visibility) | $snippet |"
        }
        Write-Output ''
    }
}
else
{
    Write-Host 'Remaining Win32 UI dependency audit'
    Write-Host ''
    foreach ($group in ($results | Group-Object Category | Sort-Object Name))
    {
        $allowedCount = @($group.Group | Where-Object { $_.Allowed }).Count
        $unallowedCount = $group.Count - $allowedCount
        Write-Host ("{0,5}  {1} ({2} allowed, {3} unallowed)" -f $group.Count, $group.Name, $allowedCount, $unallowedCount)
    }
    Write-Host ''
    $results
}

if ($FailOnFindings -and $unallowedResults.Count -gt 0)
{
    throw "Remaining Win32 UI dependency audit found $($unallowedResults.Count) unallowed finding(s)."
}
