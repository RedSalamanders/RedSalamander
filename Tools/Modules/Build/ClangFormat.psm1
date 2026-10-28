# Native formatter acquisition follows DxUi's checksum-pinned wheel extraction approach.
# The local variant preserves the previous executable on failure and verifies cached bytes.
#Requires -Version 7.0
Set-StrictMode -Version Latest
Import-Module (Join-Path $PSScriptRoot 'ArtifactOperationLock.psm1') -ErrorAction Stop

function Assert-RSClangFormatVersion {
    param([string]$Path, [string]$Version)
    $reported = & $Path --version
    if ($LASTEXITCODE -ne 0 -or ($reported -join "`n") -notmatch "^clang-format version $([regex]::Escape($Version))(?:\s|$)") {
        throw "The extracted formatter does not report clang-format ${Version}: $Path"
    }
}

function Install-RSClangFormat {
    [CmdletBinding(SupportsShouldProcess)]
    param([Parameter(Mandatory)][string]$RepositoryRoot, [string]$WheelPath = '')
    $provider = $null
    $drive = $null
    $RepositoryRoot = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($RepositoryRoot, [ref]$provider, [ref]$drive)
    if ($provider.Name -cne 'FileSystem') { throw 'Formatter repository root requires the FileSystem provider' }
    if ($WheelPath) {
        # Provider resolution follows PowerShell's location, unlike IO.Path.GetFullPath.
        # Hashing and ZIP opening must consume this one absolute, literal input.
        $WheelPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($WheelPath, [ref]$provider, [ref]$drive)
        if ($provider.Name -cne 'FileSystem') { throw 'Formatter wheel requires the FileSystem provider' }
    }
    $root = Resolve-RSGuardedRepositoryPath -RepoRoot $RepositoryRoot -GuardedRelativeRoot '.build/format' -Path '.build/format'
    $pin = Get-Content -LiteralPath (Join-Path $RepositoryRoot 'Build/clang-format.json') -Raw | ConvertFrom-Json
    if ($pin.version -cnotmatch '^\d+\.\d+\.\d+$' -or $pin.wheelSha256 -cnotmatch '^[0-9a-f]{64}$') { throw 'Invalid formatter pin' }
    $executable = Resolve-RSGuardedRepositoryPath -RepoRoot $RepositoryRoot -GuardedRelativeRoot '.build/format' -Path (Join-Path $root 'clang_format/data/bin/clang-format.exe')
    if (-not $PSCmdlet.ShouldProcess($root, "Acquire verified clang-format $($pin.version)")) { return }
    $root = Resolve-RSGuardedRepositoryPath -RepoRoot $RepositoryRoot -GuardedRelativeRoot '.build/format' -Path $root -CreateDirectory
    $cachedWheel = Resolve-RSGuardedRepositoryPath -RepoRoot $RepositoryRoot -GuardedRelativeRoot '.build/format' -Path (Join-Path $root "clang-format-$($pin.version).whl")
    $download = Resolve-RSGuardedRepositoryPath -RepoRoot $RepositoryRoot -GuardedRelativeRoot '.build/format' -Path (Join-Path $root ([guid]::NewGuid().ToString('N') + '.download'))
    $leasePath = Resolve-RSGuardedRepositoryPath -RepoRoot $RepositoryRoot -GuardedRelativeRoot '.build/format' -Path (Join-Path $root '.install.lock')
    # This shared cache is independent of native artifact profiles. Fail immediately
    # on contention; the exclusive file lease covers every guarded publication.
    $lease = [IO.File]::Open($leasePath, [IO.FileMode]::OpenOrCreate, [IO.FileAccess]::ReadWrite, [IO.FileShare]::None)
    $temporaryExecutable = $null
    try {
        if (-not $WheelPath) {
            $WheelPath = $cachedWheel
            if (-not [IO.File]::Exists($cachedWheel) -or (Get-FileHash -LiteralPath $cachedWheel -Algorithm SHA256).Hash -ine $pin.wheelSha256) {
                $index = Invoke-RestMethod -Uri 'https://pypi.org/simple/clang-format/' -Headers @{Accept='application/vnd.pypi.simple.v1+json'} -TimeoutSec 60 -ErrorAction Stop
                $candidates = @($index.files | Where-Object {$_.hashes.sha256 -ceq $pin.wheelSha256})
                if ($candidates.Count -ne 1) { throw "Package index must identify exactly one wheel matching the pinned SHA-256; found $($candidates.Count)" }
                if (([uri]$candidates[0].url).Scheme -cne 'https') { throw 'The formatter download requires HTTPS' }
                $download = Resolve-RSGuardedRepositoryPath -RepoRoot $RepositoryRoot -GuardedRelativeRoot '.build/format' -Path $download
                Invoke-WebRequest -Uri $candidates[0].url -OutFile $download -TimeoutSec 60 -ErrorAction Stop
                $download = Resolve-RSGuardedRepositoryPath -RepoRoot $RepositoryRoot -GuardedRelativeRoot '.build/format' -Path $download
                if ((Get-FileHash -LiteralPath $download -Algorithm SHA256).Hash -ine $pin.wheelSha256) { throw 'Downloaded formatter wheel does not match its pinned SHA-256' }
                $null = Publish-RSGuardedRepositoryFile -RepoRoot $RepositoryRoot -GuardedRelativeRoot '.build/format' -StagedPath $download -DestinationPath $cachedWheel
            }
        }
        # Keep the verified wheel open against writes/deletion for the entire read.
        # Hash and ZIP parsing use the same stream, not a second filesystem lookup.
        if ($WheelPath -eq $cachedWheel) {
            $WheelPath = Resolve-RSGuardedRepositoryPath -RepoRoot $RepositoryRoot -GuardedRelativeRoot '.build/format' -Path $WheelPath
        }
        $wheelStream = [IO.File]::OpenRead($WheelPath)
        try {
            $wheelHasher = [Security.Cryptography.SHA256]::Create()
            try { $wheelHash = [BitConverter]::ToString($wheelHasher.ComputeHash($wheelStream)).Replace('-', '') }
            finally { $wheelHasher.Dispose() }
            if ($wheelHash -ine $pin.wheelSha256) { throw 'Formatter wheel does not match its pinned SHA-256' }
            $wheelStream.Position = 0
            $archive = [IO.Compression.ZipArchive]::new($wheelStream, [IO.Compression.ZipArchiveMode]::Read, $true)
            try {
                $entries = @($archive.Entries | Where-Object FullName -CEQ 'clang_format/data/bin/clang-format.exe')
                if ($entries.Count -ne 1) { throw "Formatter wheel must contain exactly one native executable; found $($entries.Count)" }
                $entry = $entries[0]
                $stream = $entry.Open()
                $hasher = [Security.Cryptography.SHA256]::Create()
                try { $expectedHash = [BitConverter]::ToString($hasher.ComputeHash($stream)).Replace('-', '') }
                finally { $hasher.Dispose(); $stream.Dispose() }
                $executable = Resolve-RSGuardedRepositoryPath -RepoRoot $RepositoryRoot -GuardedRelativeRoot '.build/format' -Path $executable
                if ([IO.File]::Exists($executable) -and (Get-FileHash -LiteralPath $executable -Algorithm SHA256).Hash -ceq $expectedHash) {
                    Assert-RSClangFormatVersion $executable $pin.version
                    return $executable
                }
                $executableParent = Resolve-RSGuardedRepositoryPath -RepoRoot $RepositoryRoot -GuardedRelativeRoot '.build/format' -Path (Split-Path $executable) -CreateDirectory
                $temporaryExecutable = Resolve-RSGuardedRepositoryPath -RepoRoot $RepositoryRoot -GuardedRelativeRoot '.build/format' -Path (Join-Path $executableParent ([guid]::NewGuid().ToString('N') + '.exe'))
                [IO.Compression.ZipFileExtensions]::ExtractToFile($entry, $temporaryExecutable, $false)
                $temporaryExecutable = Resolve-RSGuardedRepositoryPath -RepoRoot $RepositoryRoot -GuardedRelativeRoot '.build/format' -Path $temporaryExecutable
                Assert-RSClangFormatVersion $temporaryExecutable $pin.version
                $null = Publish-RSGuardedRepositoryFile -RepoRoot $RepositoryRoot -GuardedRelativeRoot '.build/format' -StagedPath $temporaryExecutable -DestinationPath $executable
                } finally { $archive.Dispose() }
        } finally { $wheelStream.Dispose() }
    } finally {
        try {
            foreach ($temporary in @($download, $temporaryExecutable)) {
                if (-not $temporary) { continue }
                try { $safeTemporary = Resolve-RSGuardedRepositoryPath -RepoRoot $RepositoryRoot -GuardedRelativeRoot '.build/format' -Path $temporary }
                catch [System.Management.Automation.RuntimeException] {
                    # A changed output chain is no longer ours to clean. Preserve the
                    # original failure and leave it for guarded, non-following cleanup.
                    Write-Warning "Formatter temporary path retained because its output chain is unsafe: $temporary"
                    continue
                }
                if ([IO.File]::Exists($safeTemporary)) { [IO.File]::Delete($safeTemporary) }
            }
        } finally { $lease.Dispose() }
    }
    return $executable
}

Export-ModuleMember -Function Install-RSClangFormat
