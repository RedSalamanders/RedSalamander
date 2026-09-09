Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
. (Join-Path $repoRoot 'Tools\TestRunPlan.ps1')
$modulePath = Join-Path $repoRoot 'Tools\TerminalEngine\TerminalEngineSourceReview.psm1'
Import-Module $modulePath -Force
Import-Module (Join-Path $repoRoot 'Tools\TerminalEvidence.psm1') -Force

function Get-RSTestSha256Bytes {
    param([Parameter(Mandatory)][byte[]]$Bytes)
    $sha = [Security.Cryptography.SHA256]::Create()
    try {
        return [Convert]::ToHexString($sha.ComputeHash($Bytes)).ToLowerInvariant()
    }
    finally {
        $sha.Dispose()
    }
}

function New-RSTestSourceArchive {
    param(
        [Parameter(Mandatory)][string]$Path,
        [Parameter(Mandatory)][byte[]]$Bytes
    )

    $file = [IO.File]::Create($Path)
    try {
        $gzip = [IO.Compression.GZipStream]::new(
            $file,
            [IO.Compression.CompressionLevel]::SmallestSize,
            $true)
        try {
            $writer = [System.Formats.Tar.TarWriter]::new($gzip, $true)
            try {
                $entry = [System.Formats.Tar.PaxTarEntry]::new(
                    [System.Formats.Tar.TarEntryType]::RegularFile,
                    'root-pin/src/example.txt')
                $entry.DataStream = [IO.MemoryStream]::new($Bytes, $false)
                try {
                    $writer.WriteEntry($entry)
                }
                finally {
                    $entry.DataStream.Dispose()
                }
            }
            finally {
                $writer.Dispose()
            }
        }
        finally {
            $gzip.Dispose()
        }
    }
    finally {
        $file.Dispose()
    }
}

Describe 'TerminalEngine source-review byte proof' {
    BeforeEach {
        $script:scratch = New-RSTestSandboxScratchDirectory `
            -RepoRoot $repoRoot `
            -Harness 'terminal-engine-source-review-pester' `
            -Case ([guid]::NewGuid().ToString('N'))
        $script:sourceBytes = [Text.Encoding]::UTF8.GetBytes("alpha`r`nbeta`ngamma")
        $script:archive = Join-Path $script:scratch 'source.tar.gz'
        New-RSTestSourceArchive -Path $script:archive -Bytes $script:sourceBytes
    }

    It 'hashes LF-delimited ranges without CRLF or Unicode normalization' {
        $range = Get-RSSourceLineRangeIdentity `
            -Bytes $script:sourceBytes `
            -LineStart 1 `
            -LineEnd 2
        $expected = [Text.Encoding]::UTF8.GetBytes("alpha`r`nbeta`n")
        $range.offset | Should Be 0
        $range.sizeBytes | Should Be $expected.Length
        $range.sha256 | Should Be (Get-RSTestSha256Bytes -Bytes $expected)
    }

    It 'reads only the exact safe regular member and verifies all three digests' {
        $range = Get-RSSourceLineRangeIdentity `
            -Bytes $script:sourceBytes `
            -LineStart 2 `
            -LineEnd 3
        $blocker = [pscustomobject]@{
            source = [pscustomobject]@{
                relativePath = 'src/example.txt'
                fileSha256 = Get-RSTestSha256Bytes -Bytes $script:sourceBytes
                lineStart = 2
                lineEnd = 3
                lineRangeSha256 = [string]$range.sha256
            }
        }
        $proof = Test-RSSourceReviewBlocker `
            -ArchivePath $script:archive `
            -ExpectedArchiveSha256 (
                (Get-FileHash -LiteralPath $script:archive -Algorithm SHA256).Hash.ToLowerInvariant()) `
            -Blocker $blocker
        $proof.relativePath | Should Be 'src/example.txt'
        $proof.lineRangeSha256 | Should Be $range.sha256
        (Get-RSJcsSha256 -InputObject $proof) |
            Should Not Be '44136fa355b3678a1146ad16f7e8649e94fb4fc21fe77e8310c060f61caaff8a'
        [Text.Encoding]::UTF8.GetString(
            (ConvertTo-RSJcsUtf8Bytes -InputObject $proof)) |
            Should Match '"lineRangeSha256"'
    }

    It 'rejects archive, member, and line-range substitution' {
        $blocker = [pscustomobject]@{
            source = [pscustomobject]@{
                relativePath = 'src/example.txt'
                fileSha256 = ('0' * 64)
                lineStart = 1
                lineEnd = 1
                lineRangeSha256 = ('0' * 64)
            }
        }
        $blockerRejected = $false
        try {
            Test-RSSourceReviewBlocker `
                -ArchivePath $script:archive `
                -ExpectedArchiveSha256 (
                    (Get-FileHash -LiteralPath $script:archive -Algorithm SHA256).Hash.ToLowerInvariant()) `
                -Blocker $blocker
        }
        catch [System.Management.Automation.RuntimeException] {
            $blockerRejected = $true
        }
        $blockerRejected | Should Be $true

        $unsafeMemberRejected = $false
        try {
            Read-RSSourceArchiveMember `
                -ArchivePath $script:archive `
                -RelativePath '../example.txt'
        }
        catch [System.Management.Automation.RuntimeException] {
            $unsafeMemberRejected = $true
        }
        $unsafeMemberRejected | Should Be $true

        $invalidRangeRejected = $false
        try {
            Get-RSSourceLineRangeIdentity `
                -Bytes $script:sourceBytes `
                -LineStart 99 `
                -LineEnd 99
        }
        catch [System.Management.Automation.RuntimeException] {
            $invalidRangeRejected = $true
        }
        $invalidRangeRejected | Should Be $true
    }
}
