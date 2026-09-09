Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$script:RepositoryRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)

function Get-RSText {
    <#
    .SYNOPSIS
        Reads a repository-relative text file for a source-contract test.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    return Get-Content -LiteralPath (Join-Path $script:RepositoryRoot $Path) -Raw
}

function Assert-RSEqual {
    <#
    .SYNOPSIS
        Raises a focused test error when two scalar values differ.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [AllowNull()]
        [object]$Actual,

        [Parameter(Mandatory = $true)]
        [AllowNull()]
        [object]$Expected,

        [Parameter(Mandatory = $true)]
        [string]$Message
    )

    if ($Actual -ne $Expected) {
        throw "$Message Expected '$Expected' but got '$Actual'."
    }
}

function Assert-RSSequenceEqual {
    <#
    .SYNOPSIS
        Raises a focused test error when two ordered sequences differ.
    #>
    [CmdletBinding()]
    param(
        [AllowNull()]
        [object[]]$Actual,

        [AllowNull()]
        [object[]]$Expected,

        [Parameter(Mandatory = $true)]
        [string]$Message
    )

    $actualItems = @($Actual)
    $expectedItems = @($Expected)
    if ($actualItems.Count -ne $expectedItems.Count) {
        throw "$Message Expected $($expectedItems.Count) item(s) but got $($actualItems.Count)."
    }

    for ($index = 0; $index -lt $expectedItems.Count; $index++) {
        if ($actualItems[$index] -ne $expectedItems[$index]) {
            throw "$Message Item $index expected '$($expectedItems[$index])' but got '$($actualItems[$index])'."
        }
    }
}

function New-RSValidationGitFixture {
    <#
    .SYNOPSIS
        Creates the shared dirty/staged/untracked/delete/rename Git fixture used by
        Operation Startrail contract tests.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)][string]$Root)

    $fixtureRoot = [IO.Path]::GetFullPath($Root)
    if (Test-Path -LiteralPath $fixtureRoot) {
        throw "Validation fixture root already exists: $fixtureRoot"
    }
    [void](New-Item -ItemType Directory -Path $fixtureRoot)
    [void](New-Item -ItemType Directory -Path (Join-Path $fixtureRoot 'path with spaces'))
    [IO.File]::WriteAllText((Join-Path $fixtureRoot 'dirty.txt'), "base`n", [Text.UTF8Encoding]::new($false))
    [IO.File]::WriteAllText((Join-Path $fixtureRoot 'delete.txt'), "delete`n", [Text.UTF8Encoding]::new($false))
    [IO.File]::WriteAllText((Join-Path $fixtureRoot 'rename.txt'), "rename`n", [Text.UTF8Encoding]::new($false))
    [IO.File]::WriteAllText((Join-Path $fixtureRoot 'path with spaces\tracked.txt'), "space`n", [Text.UTF8Encoding]::new($false))

    & git -C $fixtureRoot init --quiet
    if ($LASTEXITCODE -ne 0) { throw 'Unable to initialize validation Git fixture.' }
    & git -C $fixtureRoot config user.email 'startrail-fixture@invalid.example'
    & git -C $fixtureRoot config user.name 'Startrail Fixture'
    & git -C $fixtureRoot add -- .
    & git -C $fixtureRoot commit --quiet -m 'fixture base'
    if ($LASTEXITCODE -ne 0) { throw 'Unable to commit validation Git fixture base.' }

    [IO.File]::WriteAllText((Join-Path $fixtureRoot 'dirty.txt'), "dirty`n", [Text.UTF8Encoding]::new($false))
    [IO.File]::WriteAllText((Join-Path $fixtureRoot 'staged.txt'), "staged`n", [Text.UTF8Encoding]::new($false))
    & git -C $fixtureRoot add -- staged.txt
    [IO.File]::WriteAllText((Join-Path $fixtureRoot 'untracked.txt'), "untracked`n", [Text.UTF8Encoding]::new($false))
    Remove-Item -LiteralPath (Join-Path $fixtureRoot 'delete.txt')
    & git -C $fixtureRoot mv -- rename.txt renamed.txt
    if ($LASTEXITCODE -ne 0) { throw 'Unable to create validation Git fixture rename.' }

    return [pscustomobject]@{
        Root = $fixtureRoot
        DirtyPath = 'dirty.txt'
        StagedPath = 'staged.txt'
        UntrackedPath = 'untracked.txt'
        DeletedPath = 'delete.txt'
        RenamedOldPath = 'rename.txt'
        RenamedNewPath = 'renamed.txt'
        SpacePath = 'path with spaces/tracked.txt'
    }
}

Export-ModuleMember -Function @(
    'Get-RSText',
    'Assert-RSEqual',
    'Assert-RSSequenceEqual',
    'New-RSValidationGitFixture'
)
