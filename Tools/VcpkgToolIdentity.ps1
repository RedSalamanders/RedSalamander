Set-StrictMode -Version Latest

function Get-RSVcpkgToolPin {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$RepoRoot
    )

    $pinPath = Join-Path $RepoRoot 'vcpkg-tool.json'
    if (-not (Test-Path -LiteralPath $pinPath -PathType Leaf)) {
        throw "Pinned vcpkg tool identity not found: $pinPath"
    }

    $pin = Get-Content -LiteralPath $pinPath -Raw | ConvertFrom-Json
    $repository = [string]$pin.repository
    $commit = [string]$pin.commit
    if ([string]::IsNullOrWhiteSpace($repository)) {
        throw "vcpkg-tool.json repository is missing."
    }
    if ($commit -notmatch '^[0-9a-fA-F]{40}$') {
        throw "vcpkg-tool.json commit must be a full 40-character Git commit SHA."
    }

    return [pscustomobject]@{
        Repository = $repository.Trim()
        Commit = $commit.ToLowerInvariant()
        Path = $pinPath
    }
}

function Assert-RSVcpkgToolIdentity {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$VcpkgExe,

        [Parameter(Mandatory = $true)]
        [string]$ExpectedCommit
    )

    $resolvedExe = (Resolve-Path -LiteralPath $VcpkgExe).Path
    $toolRoot = Split-Path -Parent $resolvedExe
    if (-not (Test-Path -LiteralPath (Join-Path $toolRoot '.git'))) {
        throw "vcpkg tool identity cannot be verified because '$resolvedExe' is not in a Git checkout. Use a vcpkg checkout at commit $ExpectedCommit."
    }

    $head = (& git -C $toolRoot rev-parse HEAD).Trim().ToLowerInvariant()
    if ($LASTEXITCODE -ne 0 -or $head -notmatch '^[0-9a-f]{40}$') {
        throw "Unable to read the vcpkg Git identity from '$toolRoot'."
    }
    if ($head -ne $ExpectedCommit.ToLowerInvariant()) {
        throw "vcpkg tool checkout mismatch: expected $ExpectedCommit, found $head at '$toolRoot'."
    }

    return $resolvedExe
}
