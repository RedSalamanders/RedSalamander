param(
    [string]$RepoRoot = (Resolve-Path (Join-Path $PSScriptRoot '..')).Path,
    [switch]$AsMarkdown,
    [switch]$FailOnFindings
)

$ErrorActionPreference = 'Stop'

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
        Path          = 'Common\DxUi\DxUi.ComboBox.cpp'
        LinePattern   = 'GetDC\(nullptr\)|SelectObject\(memoryDc\.get\(\), bitmap\.get\(\)\)'
        Visibility    = 'visual bitmap interop'
        Owner         = 'DxUi ComboBox'
        Reason        = 'Popup backdrop capture uses a memory DC only to snapshot pixels behind the Direct2D popup; it does not render native text, fonts, or controls.'
        ExitCondition = 'Replace when the popup backdrop path moves to DirectComposition/WIC-only capture or the backdrop effect is removed.'
    },
    @{
        Category      = 'HDC text/selection bridge'
        Path          = 'Common\DxUi\DxUi.Menu.cpp'
        LinePattern   = 'GetDC\(nullptr\)|SelectObject\(memoryDc\.get\(\), bitmap\.get\(\)\)'
        Visibility    = 'visual bitmap interop'
        Owner         = 'DxUi Menu'
        Reason        = 'Popup backdrop capture uses a memory DC only to snapshot pixels behind the Direct2D menu; it does not render native text, fonts, or controls.'
        ExitCondition = 'Replace when the popup backdrop path moves to DirectComposition/WIC-only capture or the backdrop effect is removed.'
    },
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
        Path          = 'RedSalamander\SelfTest\Commands\Commands.SelfTest.ViewCommands.cpp'
        LinePattern   = 'GetDC\(hwnd\)'
        Visibility    = 'test-only'
        Owner         = 'Commands self-tests'
        Reason        = 'Pixel probes validate rendered DxUi surfaces in deterministic command self-tests.'
        ExitCondition = 'Replace with a non-HDC render capture helper if one becomes available.'
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
        Path          = 'RedSalamander\SelfTest\Commands\Commands.SelfTest.Search.cpp'
        LinePattern   = 'CreateWindowExW\(0, L"STATIC", L"", WS_(POPUP|CHILD)'
        Visibility    = 'test-only hidden queue probe'
        Owner         = 'Find dialog command self-tests'
        Reason        = 'The queue-order test creates hidden Static probe windows without WS_VISIBLE to verify parent/child input-drain ordering.'
        ExitCondition = 'Remove if the drain-order test gains a non-HWND message queue probe.'
    },
    @{
        Category      = 'HDC text/selection bridge'
        Path          = 'Tests\DxUiTests\DxUiTests.Menu.cpp'
        LinePattern   = 'GetDC\(nullptr\)'
        Visibility    = 'test-only'
        Owner         = 'DxUiTests'
        Reason        = 'Menu test uses a screen DC only to verify popup backdrop capture behavior.'
        ExitCondition = 'Replace when popup backdrop tests use a non-HDC capture helper.'
    },
    @{
        Category      = 'Native visible control creation'
        Path          = 'Tests\DxUiTests\DxUiTestHelpers.h'
        LinePattern   = 'DxUiTestsClipboardOwner'
        Visibility    = 'test-only hidden clipboard owner'
        Owner         = 'DxUiTests'
        Reason        = 'Offscreen clipboard owner HWND exists only for deterministic clipboard tests and is not an app-owned visible control.'
        ExitCondition = 'Remove if clipboard tests move to a process-level test harness owner.'
    },
    @{
        Category      = 'HFONT handle'
        Path          = 'Tests\ViewerPETests\ViewerPETests.cpp'
        LinePattern   = 'legacyVisibleHfontSurfaceCount|HFONT text surfaces'
        Visibility    = 'test-only assertion text'
        Owner         = 'ViewerPETests'
        Reason        = 'The test references HFONT only in diagnostics and a debug counter name while asserting that visible legacy HFONT text surfaces are absent.'
        ExitCondition = 'Rename the diagnostic counter if the audit eventually excludes identifiers and string literals.'
    }
)

function Get-RelativeRepoPath
{
    param(
        [Parameter(Mandatory = $true)]
        [string]$BasePath,
        [Parameter(Mandatory = $true)]
        [string]$TargetPath
    )

    $baseFull = [IO.Path]::GetFullPath((Resolve-Path -LiteralPath $BasePath).Path)
    $targetFull = [IO.Path]::GetFullPath((Resolve-Path -LiteralPath $TargetPath).Path)

    if (! $baseFull.EndsWith([IO.Path]::DirectorySeparatorChar))
    {
        $baseFull += [IO.Path]::DirectorySeparatorChar
    }

    $baseUri = [Uri]$baseFull
    $targetUri = [Uri]$targetFull
    $relativeUri = $baseUri.MakeRelativeUri($targetUri)
    return [Uri]::UnescapeDataString($relativeUri.ToString()).Replace('/', '\')
}

function Test-IsExcludedPath
{
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    $normalized = [IO.Path]::GetFullPath($Path)
    foreach ($fragment in $excludedPathFragments)
    {
        if ($normalized.IndexOf($fragment, [StringComparison]::OrdinalIgnoreCase) -ge 0)
        {
            return $true
        }
    }

    return $false
}

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

$files = foreach ($root in $sourceRoots)
{
    $path = Join-Path $RepoRoot $root
    if (! (Test-Path -LiteralPath $path))
    {
        continue
    }

    foreach ($extension in $fileExtensions)
    {
        Get-ChildItem -LiteralPath $path -Recurse -File -Filter $extension | Where-Object {
            ! (Test-IsExcludedPath -Path $_.FullName)
        }
    }
}

$files = @($files | Sort-Object FullName -Unique)
$rawResults = foreach ($pattern in $patterns)
{
    foreach ($match in ($files | Select-String -Pattern $pattern.Pattern))
    {
        [PSCustomObject]@{
            Category   = $pattern.Category
            Path       = Get-RelativeRepoPath -BasePath $RepoRoot -TargetPath $match.Path
            LineNumber = $match.LineNumber
            Line       = $match.Line.Trim()
        }
    }
}

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
    Write-Output 'Remaining Win32 UI dependency audit'
    Write-Output ''
    foreach ($group in ($results | Group-Object Category | Sort-Object Name))
    {
        $allowedCount = @($group.Group | Where-Object { $_.Allowed }).Count
        $unallowedCount = $group.Count - $allowedCount
        Write-Output ("{0,5}  {1} ({2} allowed, {3} unallowed)" -f $group.Count, $group.Name, $allowedCount, $unallowedCount)
    }
    Write-Output ''
    $results
}

if ($FailOnFindings -and $unallowedResults.Count -gt 0)
{
    throw "Remaining Win32 UI dependency audit found $($unallowedResults.Count) unallowed finding(s)."
}
