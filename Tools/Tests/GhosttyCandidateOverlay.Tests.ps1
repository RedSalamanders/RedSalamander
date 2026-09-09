Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
$modulePath = Join-Path $repoRoot 'Tools\TerminalEngine\GhosttyCandidateOverlay.psm1'
$commandPath = Join-Path $repoRoot 'Tools\New-GhosttyTerminalRuntimeCandidateOverlayReview.ps1'
$schemaPath = Join-Path $repoRoot 'Specs\Terminal\GhosttyRuntimeCandidateOverlayReview.schema.json'

Import-Module $modulePath -Force -ErrorAction Stop

function New-RSGhosttyCandidateOverlayAbiFixture {
    $pristine = Join-Path $TestDrive 'pristine'
    $applied = Join-Path $TestDrive 'applied'
    foreach ($root in @($pristine, $applied)) {
        [void](New-Item -ItemType Directory -Path (Join-Path $root 'include\ghostty\vt') -Force)
        [void](New-Item -ItemType Directory -Path (Join-Path $root 'src\terminal\c') -Force)
    }
    @'
GHOSTTY_TERMINAL_OPT_PROGRESS_REPORT = 30,
GHOSTTY_TERMINAL_OPT_CLIPBOARD_WRITE_MAX_BYTES = 39,
'@ | Set-Content -LiteralPath (Join-Path $pristine 'include\ghostty\vt\terminal.h') -Encoding utf8NoBOM
    @'
GHOSTTY_KITTY_GRAPHICS_DATA_INVALID = 0,
GHOSTTY_KITTY_GRAPHICS_DATA_PLACEMENT_ITERATOR = 1,
GHOSTTY_KITTY_GRAPHICS_DATA_GENERATION = 2,
'@ | Set-Content -LiteralPath (Join-Path $pristine 'include\ghostty\vt\kitty_graphics.h') -Encoding utf8NoBOM
    @'
GHOSTTY_TERMINAL_OPT_PROGRESS_REPORT = 30,
GHOSTTY_TERMINAL_OPT_CLIPBOARD_WRITE_MAX_BYTES = 39,
GHOSTTY_TERMINAL_OPT_KITTY_PLACEMENT_LIMITS = 40,
GHOSTTY_TERMINAL_OPT_KITTY_CLIPBOARD_WRITE_ENABLED = 41,
'@ | Set-Content -LiteralPath (Join-Path $applied 'include\ghostty\vt\terminal.h') -Encoding utf8NoBOM
    @'
kitty_placement_limits = 40,
kitty_clipboard_write_enabled = 41,
'@ | Set-Content -LiteralPath (Join-Path $applied 'src\terminal\c\terminal.zig') -Encoding utf8NoBOM
    @'
GHOSTTY_KITTY_GRAPHICS_DATA_INVALID = 0,
GHOSTTY_KITTY_GRAPHICS_DATA_PLACEMENT_ITERATOR = 1,
GHOSTTY_KITTY_GRAPHICS_DATA_GENERATION = 2,
GHOSTTY_KITTY_GRAPHICS_DATA_PLACEMENT_COUNT = 3,
GHOSTTY_KITTY_GRAPHICS_DATA_PLACEMENT_ALLOCATED_BYTES = 4,
GHOSTTY_KITTY_GRAPHICS_DATA_PLACEMENT_COUNT_LIMIT = 5,
GHOSTTY_KITTY_GRAPHICS_DATA_PLACEMENT_BYTES_LIMIT = 6,
'@ | Set-Content -LiteralPath (Join-Path $applied 'include\ghostty\vt\kitty_graphics.h') -Encoding utf8NoBOM
    @'
placement_count = 3,
placement_allocated_bytes = 4,
placement_count_limit = 5,
placement_bytes_limit = 6,
'@ | Set-Content -LiteralPath (Join-Path $applied 'src\terminal\c\kitty_graphics.zig') -Encoding utf8NoBOM
    return [pscustomobject]@{ Pristine = $pristine; Applied = $applied }
}

Describe 'Ghostty candidate overlay review' {
    It 'proves private C and Zig values follow the complete candidate public ranges' {
        $fixture = New-RSGhosttyCandidateOverlayAbiFixture
        $abi = Get-RSGhosttyCandidateOverlayAbiReview -PristineSourcePath $fixture.Pristine -AppliedSourcePath $fixture.Applied
        $abi.publicTerminalOptionMaximum | Should Be 39
        @($abi.privateTerminalOptions).Count | Should Be 2
        (@($abi.privateTerminalOptions | ForEach-Object cName) -join ',') |
            Should Be 'GHOSTTY_TERMINAL_OPT_KITTY_PLACEMENT_LIMITS,GHOSTTY_TERMINAL_OPT_KITTY_CLIPBOARD_WRITE_ENABLED'
        (@($abi.privateTerminalOptions | ForEach-Object cValue) -join ',') | Should Be '40,41'
        $abi.publicKittyDataMaximum | Should Be 2
        @($abi.privateKittyDataValues).Count | Should Be 4
        (@($abi.privateKittyDataValues | ForEach-Object cValue) -join ',') | Should Be '3,4,5,6'
    }

    It 'rejects a private terminal option collision even when C and Zig agree' {
        $fixture = New-RSGhosttyCandidateOverlayAbiFixture
        $headerPath = Join-Path $fixture.Applied 'include\ghostty\vt\terminal.h'
        (Get-Content -LiteralPath $headerPath -Raw).Replace('PLACEMENT_LIMITS = 40', 'PLACEMENT_LIMITS = 39') |
            Set-Content -LiteralPath $headerPath -Encoding utf8NoBOM
        $zigPath = Join-Path $fixture.Applied 'src\terminal\c\terminal.zig'
        (Get-Content -LiteralPath $zigPath -Raw).Replace('= 40', '= 39') |
            Set-Content -LiteralPath $zigPath -Encoding utf8NoBOM
        $message = ''
        try {
            Get-RSGhosttyCandidateOverlayAbiReview -PristineSourcePath $fixture.Pristine -AppliedSourcePath $fixture.Applied
        }
        catch {
            $message = $_.Exception.Message
        }
        $message | Should Match 'collides with the public range through 39'
    }

    It 'rejects any private terminal option whose C and Zig values disagree' {
        $fixture = New-RSGhosttyCandidateOverlayAbiFixture
        $zigPath = Join-Path $fixture.Applied 'src\terminal\c\terminal.zig'
        (Get-Content -LiteralPath $zigPath -Raw).Replace('kitty_clipboard_write_enabled = 41', 'kitty_clipboard_write_enabled = 42') |
            Set-Content -LiteralPath $zigPath -Encoding utf8NoBOM
        $message = ''
        try {
            Get-RSGhosttyCandidateOverlayAbiReview -PristineSourcePath $fixture.Pristine -AppliedSourcePath $fixture.Applied
        }
        catch {
            $message = $_.Exception.Message
        }
        $message | Should Match "C/Zig mismatch for 'GHOSTTY_TERMINAL_OPT_KITTY_CLIPBOARD_WRITE_ENABLED': 41/42"
    }

    It 'keeps the overlay evidence schema closed and enforces review limits' {
        $schema = Get-Content -LiteralPath $schemaPath -Raw | ConvertFrom-Json -Depth 100
        $schema.properties.schemaVersion.const | Should Be 2
        $schema.additionalProperties | Should Be $false
        $schema.properties.patch.properties.fileCount.maximum | Should Be 16
        $schema.properties.concerns.maxItems | Should Be 5
        $schema.properties.patch.properties.insertions.maximum | Should Be 2500
        $schema.properties.patch.properties.deletions.maximum | Should Be 2500
        $schema.properties.abi.properties.privateTerminalOptions.maxItems | Should Be 16
    }

    It 'keeps the public command thin and rejects non-full candidate identities' {
        (Get-Content -LiteralPath $commandPath).Count | Should BeLessThan 100
        (Get-Content -LiteralPath $commandPath -Raw) | Should Match 'GhosttyCandidateOverlay\.psm1'
        $message = ''
        try {
            New-RSGhosttyRuntimeCandidateOverlayReview -RepoRoot $repoRoot -CandidateCommit main `
                -PristineSourcePath $TestDrive -AppliedSourcePath $TestDrive -PatchPath $schemaPath `
                -DescriptorPath $schemaPath -ConcernReviewPath $schemaPath
        }
        catch {
            $message = $_.Exception.Message
        }
        $message | Should Match 'lowercase full SHA'
    }
}
