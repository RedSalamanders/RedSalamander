Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
$modulePath = Join-Path $repoRoot 'Tools\TerminalEngine\GhosttyCandidateOverlay.psm1'
$commandPath = Join-Path $repoRoot 'Tools\New-GhosttyTerminalRuntimeCandidateOverlayReview.ps1'
$schemaPath = Join-Path $repoRoot 'Specs\Terminal\GhosttyRuntimeCandidateOverlayReview.schema.json'
$descriptorSchemaPath = Join-Path $repoRoot 'Specs\Terminal\GhosttyRuntimeCandidateDescriptor.schema.json'

Import-Module $modulePath -Force -ErrorAction Stop
Import-Module (Join-Path $repoRoot 'Tools/Modules/Testing/ValidationFingerprint.psm1') -ErrorAction Stop

function New-RSGhosttyCandidateOverlayAbiFixture {
    param([string]$Root = $TestDrive)
    $pristine = Join-Path $Root 'pristine'
    $applied = Join-Path $Root 'applied'
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

function New-RSGhosttyCandidateOverlayDescriptorFixture {
    # This fixture owns descriptor migration; Git transport/inventory is mocked below.
    # Real canonical readers, schema validation, publication and overlay rebinding run.
    $owner = Join-Path $TestDrive ([Guid]::NewGuid().ToString('N'))
    $candidate = 'c' * 40
    $root = Join-Path $owner ".build/TerminalEngineUpgrade/$candidate"
    [void](New-Item -ItemType Directory -Path (Join-Path $owner 'Specs/Terminal') -Force)
    [void](New-Item -ItemType Directory -Path $root -Force)
    foreach ($name in @('GhosttyRuntimeCandidateDescriptor.schema.json', 'GhosttyRuntimeCandidateOverlayReview.schema.json')) {
        Copy-Item -LiteralPath (Join-Path $repoRoot "Specs/Terminal/$name") -Destination (Join-Path $owner "Specs/Terminal/$name")
    }
    $abi = New-RSGhosttyCandidateOverlayAbiFixture -Root $root
    $binding = [ordered]@{
        sourceBinding = @{ repositoryCommit = '1' * 40 }
        canonicalLock = @{ path = 'External/TerminalEngine/GhosttyRuntimeLock.v1.json'; sha256 = 'a' * 64; commit = 'b' * 40; tree = 'b' * 40 }
        candidate = @{ repository = 'https://github.com/ghostty-org/ghostty.git'; commit = $candidate; tree = '2' * 40; version = 'fixture';
            archivePath = 'candidate.tar.gz'; archiveUrl = "https://github.com/ghostty-org/ghostty/archive/$candidate.tar.gz"; archiveSha256 = 'c' * 64 }
        ancestry = @{ aheadCount = 1; behindCount = 0 }
        observation = @{ path = 'observation.json'; sha256 = 'd' * 64 }
        review = @{ path = 'review.json'; sha256 = 'e' * 64 }
        toolchain = @{ zigVersion = '0.17.0'; archives = @(@{ hostPlatform = 'x64'; fileName = 'zig.zip'; sha256 = 'a' * 64; executableSha256 = 'b' * 64 }) }
        dependencies = @(
            @{ id = 'aro'; sourceArchiveSha256 = 'a' * 64; contentArchiveSha256 = 'b' * 64; licenseSha256 = 'c' * 64 },
            @{ id = 'uucode'; sourceArchiveSha256 = 'd' * 64; contentArchiveSha256 = 'e' * 64; licenseSha256 = 'f' * 64 })
        overlay = @{ path = 'Tools/TerminalEngine/Patches/base.patch'; normalizedTextSha256 = 'f' * 64; concernIds = @('placement') }
    }
    $descriptor = [ordered]@{ schemaVersion = 2; generatedUtc = '2026-10-29T00:00:00.000Z' }
    foreach ($key in $binding.Keys) { $descriptor[$key] = $binding[$key] }
    $descriptor.inputDigest = Get-RSCanonicalJsonDigest -InputObject $binding
    $descriptorPath = Join-Path $root 'candidate-descriptor.json'
    Write-RSAtomicCanonicalJsonFile -LiteralPath $descriptorPath -InputObject $descriptor `
        -SchemaPath (Join-Path $owner 'Specs/Terminal/GhosttyRuntimeCandidateDescriptor.schema.json') -AllowedRoot $root
    $patch = Join-Path $root 'overlay.patch'
    [IO.File]::WriteAllText($patch, "fixture overlay`n", [Text.UTF8Encoding]::new($false))
    $concerns = Join-Path $root 'concerns.json'
    @{
        candidateCommit = $candidate
        concerns = @(@{ id = 'placement'; summary = 'Fixture placement bound'; sourceEvidence = @('Screen.zig'); focusedTests = @('placement') })
        currentConcernDispositions = @(@{ id = 'placement'; disposition = 'retained'; candidateConcernIds = @('placement'); rationale = 'Retain the bounded placement contract.' })
    } | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath $concerns -Encoding utf8NoBOM
    return [pscustomobject]@{
        Root = $root; Descriptor = $descriptor; DescriptorPath = $descriptorPath
        BackupPath = Join-Path $root 'candidate-descriptor.p4.json'
        ReplacementPath = Join-Path $root 'candidate-descriptor.p4.v2.json'
        Arguments = @{ RepoRoot = $owner; CandidateCommit = $candidate; PristineSourcePath = $abi.Pristine; AppliedSourcePath = $abi.Applied;
            PatchPath = $patch; DescriptorPath = $descriptorPath; ConcernReviewPath = $concerns; GeneratedUtc = [DateTime]'2026-10-29T00:00:00Z' }
    }
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

Describe 'Ghostty candidate P4 descriptor migration' {
    BeforeEach {
        Mock -ModuleName GhosttyCandidateOverlay Read-RSGhosttyRuntimeLock {
            return [pscustomobject]@{ overlay = [pscustomobject]@{ concernIds = @('placement') } }
        }
        Mock -ModuleName GhosttyCandidateOverlay Invoke-RSGhosttyCandidateOverlayGit {
            param($RepositoryPath, $Arguments, $Context, $AllowFailure)
            $lines = [string[]]@()
            if ($Arguments[0] -eq 'rev-parse') {
                $lines = [string[]]@(if ($Arguments[1] -eq 'HEAD^{tree}') { '2' * 40 } else { 'c' * 40 })
            } elseif ($Arguments -contains '--numstat') {
                $lines = [string[]]@("1`t0`tsrc/terminal/Screen.zig")
            } elseif ($Arguments -contains '--full-index') {
                $lines = [string[]]@('fixture overlay')
            }
            return [pscustomobject]@{ ExitCode = 0; Lines = $lines }
        }
    }

    It 'rejects a legacy saved P4 before rebinding with an actionable recovery path' {
        $fixture = New-RSGhosttyCandidateOverlayDescriptorFixture
        $legacy = $fixture.Descriptor
        $legacy.schemaVersion = 1
        $legacy.Remove('dependencies')
        [IO.File]::WriteAllText($fixture.BackupPath, (ConvertTo-RSCanonicalJson $legacy) + "`n", [Text.UTF8Encoding]::new($false))
        $before = (Get-FileHash -LiteralPath $fixture.BackupPath).Hash
        $message = ''
        $arguments = $fixture.Arguments
        try { New-RSGhosttyRuntimeCandidateOverlayReview @arguments }
        catch { $message = $_.Exception.Message }
        $message | Should Match 'schema-2.*Phase4DescriptorPath'
        (Get-FileHash -LiteralPath $fixture.BackupPath).Hash | Should Be $before
        (Test-Path -LiteralPath (Join-Path $fixture.Root 'candidate-overlay-review.json')) | Should Be $false
    }

    It 'preserves the existing positional timestamp and creates the default snapshot' {
        $fixture = New-RSGhosttyCandidateOverlayDescriptorFixture
        $a = $fixture.Arguments
        $result = New-RSGhosttyRuntimeCandidateOverlayReview $a.RepoRoot $a.CandidateCommit $a.PristineSourcePath `
            $a.AppliedSourcePath $a.PatchPath $a.DescriptorPath $a.ConcernReviewPath $a.GeneratedUtc
        $result.BaseDescriptorPath | Should Be $fixture.BackupPath
        $review = Get-Content -LiteralPath $result.ReviewPath -Raw | ConvertFrom-Json -DateKind String
        $review.generatedUtc | Should Be '2026-10-29T00:00:00.000Z'
        (Test-Path -LiteralPath $fixture.BackupPath) | Should Be $true
    }

    It 'uses a regenerated schema-2 descriptor at a new P4 path while preserving the legacy evidence' {
        $fixture = New-RSGhosttyCandidateOverlayDescriptorFixture
        [IO.File]::WriteAllText($fixture.BackupPath, ('{"schemaVersion":1}' + "`n"), [Text.UTF8Encoding]::new($false))
        $before = (Get-FileHash -LiteralPath $fixture.BackupPath).Hash
        $arguments = $fixture.Arguments
        $result = New-RSGhosttyRuntimeCandidateOverlayReview @arguments -Phase4DescriptorPath $fixture.ReplacementPath
        $result.BaseDescriptorPath | Should Be $fixture.ReplacementPath
        (Get-FileHash -LiteralPath $fixture.BackupPath).Hash | Should Be $before
        $saved = Get-Content -LiteralPath $fixture.ReplacementPath -Raw | ConvertFrom-Json -DateKind String
        $updated = Get-Content -LiteralPath $fixture.DescriptorPath -Raw | ConvertFrom-Json -DateKind String
        $saved.schemaVersion | Should Be 2
        @($updated.dependencies).Count | Should Be 2
        (ConvertTo-RSCanonicalJson $updated.dependencies) | Should Be (ConvertTo-RSCanonicalJson $saved.dependencies)
        $snapshotHash = (Get-FileHash -LiteralPath $fixture.ReplacementPath).Hash
        New-RSGhosttyRuntimeCandidateOverlayReview @arguments -Phase4DescriptorPath $fixture.ReplacementPath | Out-Null
        (Get-FileHash -LiteralPath $fixture.ReplacementPath).Hash | Should Be $snapshotHash
    }

    It 'rejects a saved schema-2 descriptor belonging to another candidate tree' {
        $fixture = New-RSGhosttyCandidateOverlayDescriptorFixture
        $fixture.Descriptor.candidate.tree = '3' * 40
        Write-RSAtomicCanonicalJsonFile -LiteralPath $fixture.BackupPath -InputObject $fixture.Descriptor `
            -SchemaPath $descriptorSchemaPath -AllowedRoot $fixture.Root
        $arguments = $fixture.Arguments
        $message = ''
        try { New-RSGhosttyRuntimeCandidateOverlayReview @arguments }
        catch { $message = $_.Exception.Message }
        $message | Should Match 'candidate commit/tree does not match'
        (Test-Path -LiteralPath (Join-Path $fixture.Root 'candidate-overlay-review.json')) | Should Be $false
    }

    It 'requires a new P4 path after regenerating a different schema-2 dependency binding' {
        $fixture = New-RSGhosttyCandidateOverlayDescriptorFixture
        Copy-Item -LiteralPath $fixture.DescriptorPath -Destination $fixture.BackupPath
        $before = (Get-FileHash -LiteralPath $fixture.BackupPath).Hash
        $fixture.Descriptor.dependencies[1].licenseSha256 = '0' * 64
        $binding = [ordered]@{}
        foreach ($key in $fixture.Descriptor.Keys | Where-Object { $_ -notin @('schemaVersion', 'generatedUtc', 'inputDigest') }) {
            $binding[$key] = $fixture.Descriptor[$key]
        }
        $fixture.Descriptor.inputDigest = Get-RSCanonicalJsonDigest $binding
        Write-RSAtomicCanonicalJsonFile -LiteralPath $fixture.DescriptorPath -InputObject $fixture.Descriptor `
            -SchemaPath $descriptorSchemaPath -AllowedRoot $fixture.Root
        $arguments = $fixture.Arguments
        $message = ''
        try { New-RSGhosttyRuntimeCandidateOverlayReview @arguments }
        catch { $message = $_.Exception.Message }
        $message | Should Match 'differs from the regenerated frozen inputs.*Phase4DescriptorPath'
        New-RSGhosttyRuntimeCandidateOverlayReview @arguments -Phase4DescriptorPath $fixture.ReplacementPath | Out-Null
        $updated = Get-Content -LiteralPath $fixture.DescriptorPath -Raw | ConvertFrom-Json -DateKind String
        $updated.dependencies[1].licenseSha256 | Should Be ('0' * 64)
        (Get-FileHash -LiteralPath $fixture.BackupPath).Hash | Should Be $before
    }

    It 'rejects a changed dependency binding before creating its P4 snapshot' {
        $fixture = New-RSGhosttyCandidateOverlayDescriptorFixture
        $fixture.Descriptor.dependencies[1].licenseSha256 = '0' * 64
        Write-RSAtomicCanonicalJsonFile -LiteralPath $fixture.DescriptorPath -InputObject $fixture.Descriptor `
            -SchemaPath $descriptorSchemaPath -AllowedRoot $fixture.Root
        $arguments = $fixture.Arguments
        $message = ''
        try { New-RSGhosttyRuntimeCandidateOverlayReview @arguments }
        catch { $message = $_.Exception.Message }
        $message | Should Match 'inputDigest does not match'
        (Test-Path -LiteralPath $fixture.BackupPath) | Should Be $false
    }

    It 'rejects a P4 path that aliases the mutable descriptor' {
        $fixture = New-RSGhosttyCandidateOverlayDescriptorFixture
        $before = (Get-FileHash -LiteralPath $fixture.DescriptorPath).Hash
        $arguments = $fixture.Arguments
        $message = ''
        try { New-RSGhosttyRuntimeCandidateOverlayReview @arguments -Phase4DescriptorPath $fixture.DescriptorPath.ToUpperInvariant() }
        catch { $message = $_.Exception.Message }
        $message | Should Match 'must differ from the mutable DescriptorPath'
        (Get-FileHash -LiteralPath $fixture.DescriptorPath).Hash | Should Be $before
    }

    It 'rejects a P4 path outside its exact candidate root' {
        $fixture = New-RSGhosttyCandidateOverlayDescriptorFixture
        $outside = Join-Path $TestDrive 'outside-p4.json'
        $arguments = $fixture.Arguments
        $message = ''
        try { New-RSGhosttyRuntimeCandidateOverlayReview @arguments -Phase4DescriptorPath $outside }
        catch { $message = $_.Exception.Message }
        $message | Should Match 'Phase4DescriptorPath must remain below the exact candidate root'
        (Test-Path -LiteralPath $outside) | Should Be $false
    }
}
