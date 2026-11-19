# Prepares the complete vcpkg checkout that a pinned DxUi's vcpkg-install.ps1 restores with. The same file is in
# RedXe, RedPrism, RedTilde, RedMonitor and RedSalamander; it moves to the shared toolchain home with CV-5-07 and
# is unnecessary once every DxUi pin holds DxUi's own complete-clone installer. ASCII only; parses in Windows
# PowerShell 5.1.
Set-StrictMode -Version Latest

function Initialize-VcpkgToolCheckout {
    <#
    .SYNOPSIS
    Makes ToolRoot a complete, clean vcpkg checkout at the commit a vcpkg-tool.json names.
    .DESCRIPTION
    A DxUi installer pinned before 2026-11-19 clones vcpkg with --filter=blob:none unless <OutputRoot>\vcpkg-tool
    already is a clean checkout of the repository and commit its vcpkg-tool.json names, which it then reuses. vcpkg
    extracts each versioned port with git read-tree and checkout-index in an environment that holds only
    GIT_INDEX_FILE. Without SystemRoot that Git resolves no host, so a blob a partial clone left on the server fails
    the restore with "Could not resolve host: github.com". Call this with that installer's <OutputRoot>\vcpkg-tool
    before running it: it clones completely, or completes an earlier partial clone in place, and checks out the
    pinned commit. It never deletes anything and refuses a checkout of another repository.
    .PARAMETER ToolIdentityPath
    The vcpkg-tool.json that names the vcpkg repository URL and full commit.
    .PARAMETER ToolRoot
    The managed vcpkg checkout directory, for example <OutputRoot>\vcpkg-tool.
    .OUTPUTS
    None. Git progress goes to the host; failures throw.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$ToolIdentityPath, [Parameter(Mandatory)][string]$ToolRoot)
    $identity = Get-Content -Raw -LiteralPath $ToolIdentityPath | ConvertFrom-Json
    $repository = [string]$identity.repository
    $commit = ([string]$identity.commit).ToLowerInvariant()
    if ([string]::IsNullOrWhiteSpace($repository) -or $commit -notmatch '^[0-9a-f]{40}$') {
        throw "$ToolIdentityPath must contain a repository URL and a full 40-character commit hash."
    }
    $root = [IO.Path]::GetFullPath($ToolRoot)
    if (-not (Test-Path -LiteralPath (Join-Path $root '.git') -PathType Container)) {
        if (Test-Path -LiteralPath $root) { throw "The managed vcpkg path exists but is not a Git checkout: $root" }
        New-Item -ItemType Directory -Force -Path (Split-Path -Parent $root) | Out-Null
        Write-Host "Cloning the complete vcpkg checkout into $root" -ForegroundColor Cyan
        # Set only this managed checkout: deep consumer output roots may exceed the traditional Win32 path limit.
        & git.exe clone -c core.longpaths=true --no-checkout $repository $root
        if ($LASTEXITCODE -ne 0) { throw "git clone failed with exit code $LASTEXITCODE." }
    } else {
        $origin = & git.exe -C $root remote get-url origin
        if ($LASTEXITCODE -ne 0 -or ([string]$origin).TrimEnd('/') -ne $repository.TrimEnd('/')) {
            throw "The managed vcpkg checkout does not use the pinned repository: $repository"
        }
        # The filter setting goes only once every object is local, so an interrupted completion is retried.
        if (& git.exe -C $root config --get remote.origin.partialclonefilter) {
            Write-Host "Completing the partial vcpkg clone in $root" -ForegroundColor Cyan
            if ((& git.exe -C $root rev-parse --is-shallow-repository) -eq 'true') {
                & git.exe -C $root fetch --unshallow origin
                if ($LASTEXITCODE -ne 0) { throw "Unable to unshallow the managed vcpkg checkout (exit code $LASTEXITCODE)." }
            }
            & git.exe -C $root fetch --refetch --no-filter origin
            if ($LASTEXITCODE -ne 0) { throw "Unable to complete the managed vcpkg checkout (exit code $LASTEXITCODE)." }
            & git.exe -C $root config --unset-all remote.origin.partialclonefilter
            & git.exe -C $root config --unset-all remote.origin.promisor
        }
    }
    if ((& git.exe -C $root rev-parse --verify --quiet HEAD) -ne $commit) {
        # The default branch normally holds the pinned commit; fetch it, with its history, only when it is missing.
        if (-not (& git.exe -C $root rev-parse --verify --quiet "$commit^{commit}")) {
            & git.exe -C $root fetch origin $commit
            if ($LASTEXITCODE -ne 0) { throw "Unable to fetch pinned vcpkg commit $commit." }
        }
        & git.exe -C $root checkout --detach --quiet $commit
        if ($LASTEXITCODE -ne 0) { throw "Unable to check out pinned vcpkg commit $commit." }
    }
    $global:LASTEXITCODE = 0
}

Export-ModuleMember -Function Initialize-VcpkgToolCheckout
