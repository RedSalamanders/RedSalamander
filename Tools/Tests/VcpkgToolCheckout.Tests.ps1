Import-Module (Join-Path $PSScriptRoot 'TestSupport.psm1') -Force -ErrorAction Stop
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
Import-Module (Join-Path $repoRoot 'Tools/Modules/Build/VcpkgToolCheckout.psm1') -Force -ErrorAction Stop
Import-Module (Join-Path $repoRoot 'Tools/Modules/Build/ProcessStreaming.psm1') -ErrorAction Stop
Import-Module (Join-Path $repoRoot 'Tools/Modules/Testing/TestSandbox.psm1') -ErrorAction Stop

function Invoke-VcpkgCheckoutFixtureGit {
    param([string]$Root, [string[]]$Arguments)
    $output = @(Invoke-RSTestFixtureGit -Arguments @('-C'; $Root; $Arguments) 2>&1)
    if ($LASTEXITCODE -ne 0) { throw ($output -join [Environment]::NewLine) }
    return (($output | ForEach-Object { [string]$_ }) -join [Environment]::NewLine).Trim()
}

function New-VcpkgCheckoutFixture {
    # A registry served over file:// with partial-clone filtering, like GitHub serves microsoft/vcpkg. The pinned commit is
    # the third of five, and the second commit's tree holds a blob that neither the pinned commit nor the branch tip has:
    # an older versioned port that a partial clone of the branch never fetched.
    $root = New-RSTestSandboxScratchDirectory -RepoRoot $repoRoot -Harness tools-pester -Case ('vcpkg-checkout-' + [Guid]::NewGuid().ToString('N'))
    $registry = Join-Path $root 'registry'
    [void][IO.Directory]::CreateDirectory($registry)
    $null = Invoke-VcpkgCheckoutFixtureGit $registry @('init', '--quiet', '--initial-branch=master')
    $null = Invoke-VcpkgCheckoutFixtureGit $registry @('config', 'uploadpack.allowFilter', 'true')
    $null = Invoke-VcpkgCheckoutFixtureGit $registry @('config', 'uploadpack.allowAnySHA1InWant', 'true')
    $commits = [Collections.Generic.List[string]]::new()
    foreach ($index in 1..5) {
        [IO.File]::WriteAllText((Join-Path $registry "port$index.txt"), "port $index")
        [IO.File]::WriteAllText((Join-Path $registry 'versions.txt'), "versions $index")
        $null = Invoke-VcpkgCheckoutFixtureGit $registry @('add', '.')
        $null = Invoke-VcpkgCheckoutFixtureGit $registry @('-c', 'user.name=CI Fixture', '-c', 'user.email=ci@example.invalid', 'commit', '--quiet', '-m', "c$index")
        $commits.Add((Invoke-VcpkgCheckoutFixtureGit $registry @('rev-parse', 'HEAD')))
    }
    $url = 'file:///' + $registry.Replace('\', '/')
    $identity = Join-Path $root 'vcpkg-tool.json'
    [IO.File]::WriteAllText($identity, (@{ repository = $url; commit = $commits[2] } | ConvertTo-Json))
    return @{
        Root = $root; Url = $url; Identity = $identity; Pinned = $commits[2]
        OlderTree = (Invoke-VcpkgCheckoutFixtureGit $registry @('rev-parse', "$($commits[1])^{tree}"))
    }
}

function Test-VcpkgOfflineTreeExtraction {
    # vcpkg's own extraction: read-tree into a private index, then checkout-index. GIT_NO_LAZY_FETCH makes Git refuse any
    # promisor fetch, which is what vcpkg's environment (GIT_INDEX_FILE only, no SystemRoot) amounts to on a real host.
    param([string]$ToolRoot, [string]$Tree, [string]$Scratch)
    $target = Join-Path $Scratch ('extract-' + [Guid]::NewGuid().ToString('N').Substring(0, 8))
    [void][IO.Directory]::CreateDirectory($target)
    $environment = @{ GIT_INDEX_FILE = (Join-Path $Scratch ([Guid]::NewGuid().ToString('N') + '.index')); GIT_NO_LAZY_FETCH = '1' }
    $git = (Get-Command git -CommandType Application -ErrorAction Stop | Select-Object -First 1).Source
    $gitDir = Join-Path $ToolRoot '.git'
    $quiet = { param($Line, $IsError) }
    $read = Invoke-RSStreamingProcess -FilePath $git -Arguments @('--git-dir', $gitDir, 'read-tree', $Tree) `
        -WorkingDirectory $Scratch -AdditionalEnvironment $environment -OutputLineCallback $quiet
    if ($read -ne 0) { return $false }
    $checkout = Invoke-RSStreamingProcess -FilePath $git -Arguments @('--git-dir', $gitDir, '--work-tree', $target, 'checkout-index', '-af', "--prefix=$target/") `
        -WorkingDirectory $Scratch -AdditionalEnvironment $environment -OutputLineCallback $quiet
    return ($checkout -eq 0 -and (Test-Path -LiteralPath (Join-Path $target 'versions.txt') -PathType Leaf))
}

function Assert-VcpkgCheckoutReady {
    param([hashtable]$Fixture, [string]$ToolRoot)
    (Invoke-VcpkgCheckoutFixtureGit $ToolRoot @('remote', 'get-url', 'origin')) | Should Be $Fixture.Url
    (Invoke-VcpkgCheckoutFixtureGit $ToolRoot @('rev-parse', 'HEAD')) | Should Be $Fixture.Pinned
    (Invoke-VcpkgCheckoutFixtureGit $ToolRoot @('status', '--porcelain')) | Should Be ''
    (Invoke-VcpkgCheckoutFixtureGit $ToolRoot @('rev-parse', '--is-shallow-repository')) | Should Be 'false'
    (Invoke-VcpkgCheckoutFixtureGit $ToolRoot @('config', '--list', '--local')) | Should Not Match 'partialclonefilter'
    (Test-VcpkgOfflineTreeExtraction -ToolRoot $ToolRoot -Tree $Fixture.OlderTree -Scratch $Fixture.Root) | Should Be $true
}

function Get-VcpkgCheckoutFailure {
    param([scriptblock]$Action)
    try { $null = & $Action } catch { return $_.Exception.Message }
    throw 'Expected the vcpkg checkout preparation to refuse the fixture.'
}

Describe 'Complete vcpkg checkout for the pinned DxUi restore' {
    It 'clones completely at the pinned commit, so vcpkg extracts an older port tree offline, and is idempotent' {
        $fixture = New-VcpkgCheckoutFixture
        $tool = Join-Path $fixture.Root 'output/vcpkg-tool'
        $result = Invoke-RSTestBlock -ScriptBlock { Initialize-VcpkgToolCheckout -ToolIdentityPath $fixture.Identity -ToolRoot $tool }
        $result.Succeeded | Should Be $true
        Assert-VcpkgCheckoutReady -Fixture $fixture -ToolRoot $tool
        $again = Invoke-RSTestBlock -ScriptBlock { Initialize-VcpkgToolCheckout -ToolIdentityPath $fixture.Identity -ToolRoot $tool }
        $again.Succeeded | Should Be $true
        Assert-VcpkgCheckoutReady -Fixture $fixture -ToolRoot $tool
    }

    It 'completes in place the partial, shallow clone an earlier pinned installer left' {
        $fixture = New-VcpkgCheckoutFixture
        $output = Join-Path $fixture.Root 'output'
        [void][IO.Directory]::CreateDirectory($output)
        $tool = Join-Path $output 'vcpkg-tool'
        # What the pinned DxUi installer leaves: a blob-filtered clone, then a depth-1 fetch and checkout of the pin.
        $null = Invoke-VcpkgCheckoutFixtureGit $output @('clone', '--quiet', '--filter=blob:none', $fixture.Url, $tool)
        $null = Invoke-VcpkgCheckoutFixtureGit $tool @('fetch', '--quiet', '--depth', '1', 'origin', $fixture.Pinned)
        $null = Invoke-VcpkgCheckoutFixtureGit $tool @('checkout', '--quiet', '--detach', $fixture.Pinned)
        (Invoke-VcpkgCheckoutFixtureGit $tool @('rev-parse', '--is-shallow-repository')) | Should Be 'true'
        # Negative control: this state is exactly what fails vcpkg.
        (Test-VcpkgOfflineTreeExtraction -ToolRoot $tool -Tree $fixture.OlderTree -Scratch $fixture.Root) | Should Be $false
        $result = Invoke-RSTestBlock -ScriptBlock { Initialize-VcpkgToolCheckout -ToolIdentityPath $fixture.Identity -ToolRoot $tool }
        $result.Succeeded | Should Be $true
        Assert-VcpkgCheckoutReady -Fixture $fixture -ToolRoot $tool
    }

    It 'refuses another repository, a directory that is not a checkout and an incomplete identity, without deleting anything' {
        $fixture = New-VcpkgCheckoutFixture
        $foreign = Join-Path $fixture.Root 'foreign/vcpkg-tool'
        [void][IO.Directory]::CreateDirectory((Split-Path -Parent $foreign))
        $null = Invoke-VcpkgCheckoutFixtureGit (Split-Path -Parent $foreign) @('clone', '--quiet', $fixture.Url, $foreign)
        $null = Invoke-VcpkgCheckoutFixtureGit $foreign @('remote', 'set-url', 'origin', 'https://example.invalid/other')
        (Get-VcpkgCheckoutFailure { Initialize-VcpkgToolCheckout -ToolIdentityPath $fixture.Identity -ToolRoot $foreign }) |
            Should Match 'does not use the pinned repository'
        (Invoke-VcpkgCheckoutFixtureGit $foreign @('remote', 'get-url', 'origin')) | Should Be 'https://example.invalid/other'
        $plain = Join-Path $fixture.Root 'plain/vcpkg-tool'
        [void][IO.Directory]::CreateDirectory($plain)
        [IO.File]::WriteAllText((Join-Path $plain 'kept.txt'), 'kept')
        (Get-VcpkgCheckoutFailure { Initialize-VcpkgToolCheckout -ToolIdentityPath $fixture.Identity -ToolRoot $plain }) |
            Should Match 'exists but is not a Git checkout'
        Test-Path -LiteralPath (Join-Path $plain 'kept.txt') -PathType Leaf | Should Be $true
        $short = Join-Path $fixture.Root 'short.json'
        [IO.File]::WriteAllText($short, '{"repository":"https://github.com/microsoft/vcpkg","commit":"abc"}')
        (Get-VcpkgCheckoutFailure { Initialize-VcpkgToolCheckout -ToolIdentityPath $short -ToolRoot (Join-Path $fixture.Root 'unused/vcpkg-tool') }) |
            Should Match 'full 40-character commit hash'
        Test-Path -LiteralPath (Join-Path $fixture.Root 'unused') | Should Be $false
    }

    It 'prepares the checkout in the contained restore process immediately before the pinned installer, with literal quoted paths' {
        $text = [IO.File]::ReadAllText((Join-Path $repoRoot 'Tools/Modules/Build/DxUiDependency.psm1'))
        $find = {
            param([string]$Source)
            $tokens = $null; $errors = $null
            $ast = [Management.Automation.Language.Parser]::ParseInput($Source, [ref]$tokens, [ref]$errors)
            if ($errors.Count) { return $null }
            return $ast.Find({ param($node) $node -is [Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -ceq 'Restore-RSDxUiDependency' }, $true)
        }
        $valid = {
            param([string]$Source)
            $restore = & $find $Source
            if ($null -eq $restore) { return $false }
            $calls = @($restore.FindAll({ param($node) $node -is [Management.Automation.Language.CommandAst] -and
                        $node.GetCommandName() -ceq 'Invoke-RSDxUiRestoreProcess' }, $true) | Sort-Object { $_.Extent.StartOffset })
            if ($calls.Count -ne 2) { return $false }
            return $calls[0].Extent.Text.Contains("'-Command',`$checkout") -and $calls[1].Extent.Text.Contains("'vcpkg-install.ps1'") -and
                $restore.Extent.Text.Contains("Initialize-VcpkgToolCheckout -ToolIdentityPath {1} -ToolRoot {2}")
        }
        (& $valid $text) | Should Be $true
        $without = $text.Replace("Invoke-RSDxUiRestoreProcess `$RepoRoot `$pwshExe @('-NoProfile','-ExecutionPolicy','Bypass','-Command',`$checkout)", '')
        ($without -cne $text) | Should Be $true
        (& $valid $without) | Should Be $false

        # Evaluate the module's own composition with paths that need quoting, and run it the way the restore does.
        $restore = & $find $text
        $assignments = @($restore.FindAll({ param($node) $node -is [Management.Automation.Language.AssignmentStatementAst] -and
                    $node.Left.Extent.Text -in @('$quote', '$checkout') }, $true) | Sort-Object { $_.Extent.StartOffset })
        $assignments.Count | Should Be 2
        $statements = ($assignments | ForEach-Object { $_.Extent.Text }) -join [Environment]::NewLine
        $compose = [scriptblock]::Create("param(`$moduleRoot, `$source, `$output)`n" + $statements.Replace('$PSScriptRoot', '$moduleRoot') + "`n`$checkout")
        $fixture = New-VcpkgCheckoutFixture
        $source = Join-Path $fixture.Root "pinned DxUi's source"
        $output = Join-Path $fixture.Root "output root's 1"
        foreach ($directory in @($source, $output)) { [void][IO.Directory]::CreateDirectory($directory) }
        Copy-Item -LiteralPath $fixture.Identity -Destination (Join-Path $source 'vcpkg-tool.json')
        $command = & $compose (Join-Path $repoRoot 'Tools/Modules/Build') $source $output
        $shell = (Get-Process -Id $PID).Path
        $lines = [Collections.Generic.List[string]]::new()
        $exit = Invoke-RSStreamingProcess -FilePath $shell -Arguments @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-Command', $command) `
            -WorkingDirectory $fixture.Root -OutputLineCallback { param($Line, $IsError) $lines.Add([string]$Line) }.GetNewClosure()
        $exit | Should Be 0
        Assert-VcpkgCheckoutReady -Fixture $fixture -ToolRoot (Join-Path $output 'vcpkg-tool')
    }
}
