Import-Module (Join-Path $PSScriptRoot 'TestSupport.psm1') -Force -ErrorAction Stop
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
Import-Module (Join-Path $repoRoot 'Tools/Modules/Testing/TestRunPlan.psm1') -Force
Import-Module (Join-Path $repoRoot 'Tools/Modules/Build/VcpkgInstallSafety.psm1') -Force

function Invoke-BootstrapFixtureGit {
    param([string]$Root,[string[]]$Arguments)
    $output = @(Invoke-RSTestFixtureGit -Arguments @('-C'; $Root; $Arguments) 2>&1)
    if ($LASTEXITCODE -ne 0) { throw ($output -join [Environment]::NewLine) }
    return ($output -join [Environment]::NewLine)
}

function New-VcpkgBootstrapFixture {
    $root = New-RSTestSandboxScratchDirectory -RepoRoot $repoRoot -Harness tools-pester -Case ('vcpkg-bootstrap-' + [Guid]::NewGuid().ToString('N'))
    $source = Join-Path $root 'source'
    $consumer = Join-Path $root 'consumer'
    foreach ($path in @('source/scripts','source/versions','source/ports','consumer')) {
        [void][IO.Directory]::CreateDirectory((Join-Path $root $path))
    }
    [IO.File]::WriteAllText((Join-Path $source 'bootstrap-vcpkg.bat'), ('@echo off','echo fixture>"%~dp0vcpkg.exe"','exit /b 0' -join [Environment]::NewLine))
    [IO.File]::WriteAllText((Join-Path $source '.gitignore'),'vcpkg.exe')
    [IO.File]::WriteAllText((Join-Path $source 'scripts/vcpkg-tools.json'),'{"schema-version":1}')
    [IO.File]::WriteAllText((Join-Path $source 'versions/baseline.json'),'{"default":{}}')
    $null = Invoke-BootstrapFixtureGit $source @('init')
    $null = Invoke-BootstrapFixtureGit $source @('add','.')
    $null = Invoke-BootstrapFixtureGit $source @('-c','user.name=CI Fixture','-c','user.email=ci@example.invalid','commit','-m','baseline')
    $baseline = Invoke-BootstrapFixtureGit $source @('rev-parse','HEAD')
    [IO.File]::WriteAllText((Join-Path $source 'scripts/vcpkg-tools.json'),'{"schema-version":2}')
    $null = Invoke-BootstrapFixtureGit $source @('add','.')
    $null = Invoke-BootstrapFixtureGit $source @('-c','user.name=CI Fixture','-c','user.email=ci@example.invalid','commit','-m','tool')
    $tool = Invoke-BootstrapFixtureGit $source @('rev-parse','HEAD')
    [IO.File]::WriteAllText((Join-Path $consumer 'vcpkg.json'),(@{ 'builtin-baseline'=$baseline; dependencies=@() }|ConvertTo-Json))
    [IO.File]::WriteAllText((Join-Path $consumer 'vcpkg-tool.json'),(@{ repository=$source; commit=$tool }|ConvertTo-Json))
    return @{ Root=$root; Consumer=$consumer; Source=$source; Baseline=$baseline; Tool=$tool }
}

function Get-BootstrapFailure {
    param([scriptblock]$Action)
    try { $null = & $Action } catch { return $_.Exception.Message }
    throw 'Expected pinned bootstrap to refuse the fixture.'
}

Describe 'Pinned cold-runner vcpkg provisioning' {
    BeforeAll {
        $script:bootstrapModule = Get-Module VcpkgInstallSafety
        $script:originalBootstrapPayload = & $script:bootstrapModule { (Get-Command Invoke-RSVcpkgBootstrapPayload).ScriptBlock }
        $native = Get-Command Invoke-RSTestFixtureNative
        $newContext = Get-Command New-RSTestRunContext
        $contained = Get-Command Assert-RSTestSandboxContainedPath
        $repository = $repoRoot
        $launcher = {
            param([string]$BootstrapPath)
            $context = & $newContext -RepoRoot $repository
            $owned = & $contained -TestRoot $context.TestRoot -Path $BootstrapPath -RequireDescendant
            if ([IO.Path]::GetExtension($owned) -ine '.bat' -or $owned.IndexOfAny([char[]]'"%&|<>^!') -ge 0 -or $owned.Contains("`r") -or $owned.Contains("`n")) {
                throw 'The owned batch fixture requires a literal safe path.'
            }
            # CALL makes a spaced batch path a normal quoted argument after /c;
            # the real batch executes inside this explicitly no-window cmd process.
            & $native -FilePath $env:ComSpec -Arguments @('/d','/s','/c','call',$owned,'-disableMetrics') -WorkingDirectory (Get-Location).ProviderPath | Out-Host
            return [int]$global:LASTEXITCODE
        }.GetNewClosure()
        & $script:bootstrapModule { param($invoke) $script:RSBootstrapFixtureLauncher = $invoke } $launcher
        Mock Invoke-RSVcpkgBootstrapPayload -ModuleName VcpkgInstallSafety {
            param($BootstrapPath)
            return (& $script:RSBootstrapFixtureLauncher $BootstrapPath)
        }
    }

    AfterAll {
        & $script:bootstrapModule {
            param($original)
            Set-Item -LiteralPath Function:script:Invoke-RSVcpkgBootstrapPayload -Value $original
            Remove-Variable -Name RSBootstrapFixtureLauncher -Scope Script -ErrorAction SilentlyContinue
        } $script:originalBootstrapPayload
    }

    It 'preserves the production inherited payload and real initializer exit checks' {
        $source = Get-Content -LiteralPath (Join-Path $repoRoot 'Tools/Modules/Build/VcpkgInstallSafety.psm1') -Raw
        $tokens = $null; $errors = $null
        $ast = [Management.Automation.Language.Parser]::ParseInput($source, [ref]$tokens, [ref]$errors)
        $errors.Count | Should Be 0
        $payload = $ast.Find({param($node) $node -is [Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -ceq 'Invoke-RSVcpkgBootstrapPayload'}, $true)
        $payload.Extent.Text | Should Match '& \$BootstrapPath -disableMetrics \| Out-Host'
        $payload.Extent.Text | Should Match 'return \[int\]\$LASTEXITCODE'
        $payload.Extent.Text | Should Not Match 'Invoke-RSTestFixture|Invoke-RSCapturedProcess|Invoke-RSStreamingProcess|cmd\.exe'
        $source | Should Match '\$bootstrapExit = Invoke-RSVcpkgBootstrapPayload -BootstrapPath \$bootstrap'
        $source | Should Match 'if \(\$bootstrapExit -ne 0\) \{ throw "Pinned vcpkg bootstrap failed with exit code \$bootstrapExit\." \}'
    }

    It 'executes an owned spaced batch with actual arguments marker and exit 19' {
        $root = New-RSTestSandboxScratchDirectory -RepoRoot $repoRoot -Harness tools-pester -Case ('quoted-bootstrap-' + [Guid]::NewGuid().ToString('N'))
        $directory = Join-Path $root 'owned batch with spaces'
        [void][IO.Directory]::CreateDirectory($directory)
        $batch = Join-Path $directory 'bootstrap-vcpkg.bat'
        [IO.File]::WriteAllLines($batch, @('@echo off','echo %~1>"%~dp0received.args"','echo real-payload>"%~dp0marker.txt"','exit /b 19'))
        $code = & $script:bootstrapModule { param($path) Invoke-RSVcpkgBootstrapPayload -BootstrapPath $path } $batch
        $code | Should Be 19
        [IO.File]::ReadAllText((Join-Path $directory 'received.args')).Trim() | Should Be '-disableMetrics'
        [IO.File]::ReadAllText((Join-Path $directory 'marker.txt')).Trim() | Should Be 'real-payload'
    }

    It 'restores separate exact tool and baseline identities while preserving download-only cache state' {
        $fixture = New-VcpkgBootstrapFixture
        $downloads = Join-Path $fixture.Consumer '.build/vcpkg-tool/downloads'
        [void][IO.Directory]::CreateDirectory($downloads)
        [IO.File]::WriteAllText((Join-Path $downloads 'archive.zip'),'preserved download')
        $context = Initialize-RSVcpkgToolContext -RepoRoot $fixture.Consumer
        (Invoke-BootstrapFixtureGit $context.ToolRoot @('rev-parse','HEAD')) | Should Be $fixture.Tool
        (Invoke-BootstrapFixtureGit $context.RegistryRoot @('rev-parse','HEAD')) | Should Be $fixture.Baseline
        [IO.File]::Exists($context.Executable) | Should Be $true
        [IO.File]::ReadAllText((Join-Path $context.ToolRoot 'downloads/archive.zip')) | Should Be 'preserved download'
        (Initialize-RSVcpkgToolContext -RepoRoot $fixture.Consumer).ToolCommit | Should Be $fixture.Tool
    }

    It 'refuses dirty managed checkouts without resetting or deleting their contents' {
        $fixture = New-VcpkgBootstrapFixture
        $context = Initialize-RSVcpkgToolContext -RepoRoot $fixture.Consumer
        $metadata = Join-Path $context.ToolRoot 'scripts/vcpkg-tools.json'
        [IO.File]::WriteAllText($metadata,'{"schema-version":99}')
        Get-BootstrapFailure { Initialize-RSVcpkgToolContext -RepoRoot $fixture.Consumer } | Should Match 'dirty'
        [IO.File]::ReadAllText($metadata) | Should Be '{"schema-version":99}'
    }

    It 'rejects a tool older than the selected ports schema before bootstrapping' {
        $fixture = New-VcpkgBootstrapFixture
        [IO.File]::WriteAllText((Join-Path $fixture.Consumer 'vcpkg.json'),(@{ 'builtin-baseline'=$fixture.Tool }|ConvertTo-Json))
        [IO.File]::WriteAllText((Join-Path $fixture.Consumer 'vcpkg-tool.json'),(@{ repository=$fixture.Source; commit=$fixture.Baseline }|ConvertTo-Json))
        Get-BootstrapFailure { Initialize-RSVcpkgToolContext -RepoRoot $fixture.Consumer } | Should Match 'schema is incompatible'
        Test-Path -LiteralPath (Join-Path $fixture.Consumer '.build/vcpkg-tool/vcpkg.exe') | Should Be $false
    }

    It 'propagates an actual failing bootstrap exit and never fabricates an executable' {
        $fixture = New-VcpkgBootstrapFixture
        [IO.File]::WriteAllText((Join-Path $fixture.Source 'bootstrap-vcpkg.bat'),('@echo off','exit /b 19' -join [Environment]::NewLine))
        $null = Invoke-BootstrapFixtureGit $fixture.Source @('add','.')
        $null = Invoke-BootstrapFixtureGit $fixture.Source @('-c','user.name=CI Fixture','-c','user.email=ci@example.invalid','commit','-m','broken bootstrap')
        $commit = Invoke-BootstrapFixtureGit $fixture.Source @('rev-parse','HEAD')
        [IO.File]::WriteAllText((Join-Path $fixture.Consumer 'vcpkg-tool.json'),(@{ repository=$fixture.Source; commit=$commit }|ConvertTo-Json))
        Get-BootstrapFailure { Initialize-RSVcpkgToolContext -RepoRoot $fixture.Consumer } | Should Match 'exit code 19'
        Test-Path -LiteralPath (Join-Path $fixture.Consumer '.build/vcpkg-tool/vcpkg.exe') | Should Be $false
    }

    It 'rejects malformed pins before creating build output' {
        $fixture = New-VcpkgBootstrapFixture
        [IO.File]::WriteAllText((Join-Path $fixture.Consumer 'vcpkg-tool.json'),'{"repository":"-unsafe","commit":"short"}')
        Get-BootstrapFailure { Initialize-RSVcpkgToolContext -RepoRoot $fixture.Consumer } | Should Match 'full tool/baseline'
        Test-Path -LiteralPath (Join-Path $fixture.Consumer '.build') | Should Be $false
    }

    It 'selects the matching platform and Debug or Release libraries without changing the parent environment' {
        $fixture = New-VcpkgBootstrapFixture
        $tool = Join-Path $fixture.Consumer '.build/vcpkg-tool'
        $integration = Join-Path $tool 'scripts/buildsystems/msbuild'
        [void][IO.Directory]::CreateDirectory($integration)
        foreach ($file in @('vcpkg.props','vcpkg.targets')) { [IO.File]::WriteAllText((Join-Path $integration $file),'<Project/>') }
        $oldCL = $env:CL
        foreach ($platform in @('x64','ARM64')) {
            $triplet = if ($platform -eq 'ARM64') { 'arm64-windows' } else { 'x64-windows' }
            $installed = Join-Path $fixture.Consumer ('.build/vcpkg_installed/' + $platform.ToLowerInvariant() + '/' + $triplet)
            foreach ($path in @('include','lib/manual-link','bin','debug/lib/manual-link','debug/bin')) {
                [void][IO.Directory]::CreateDirectory((Join-Path $installed $path))
            }
            foreach ($configuration in @('Release','Debug','ASan Debug')) {
                $environment = Get-RSVcpkgBuildEnvironment -RepoRoot $fixture.Consumer -ToolRoot $tool -Platform $platform -Configuration $configuration
                $expected = if ($configuration -like '*Debug*') { 'Debug' } else { 'Release' }
                $environment.VcpkgConfiguration | Should Be $expected
                $environment.VCPKG_DEFAULT_TRIPLET | Should Be $triplet
                $environment.ForceImportAfterCppTargets | Should Be (Join-Path $integration 'vcpkg.targets')
                $environment.VcpkgManifestInstall | Should Be 'false'
                $environment.LINK | Should Match 'manual-link'
                $env:CL | Should Be $oldCL
            }
        }
        Remove-Item -LiteralPath (Join-Path $integration 'vcpkg.targets')
        Get-BootstrapFailure { Get-RSVcpkgBuildEnvironment -RepoRoot $fixture.Consumer -ToolRoot $tool -Platform x64 -Configuration Release } | Should Match 'build input not found'
    }

    It 'keeps canonical compiler paths in the child without importing ambient CL or LINK flags' {
        Import-Module (Join-Path $repoRoot 'Tools/Modules/Build/SanitizedEnvironment.psm1') -ErrorAction Stop
        $root = New-RSTestSandboxScratchDirectory -RepoRoot $repoRoot -Harness tools-pester -Case ('vcpkg-child-environment-' + [Guid]::NewGuid().ToString('N'))
        $tool = Join-Path $root '.build/vcpkg-tool'
        $integration = Join-Path $tool 'scripts/buildsystems/msbuild'
        $installed = Join-Path $root '.build/vcpkg_installed/x64/x64-windows'
        foreach ($directory in @($integration,(Join-Path $installed 'include'),(Join-Path $installed 'lib/manual-link'))) {
            [void][IO.Directory]::CreateDirectory($directory)
        }
        foreach ($file in @('vcpkg.props','vcpkg.targets')) { [IO.File]::WriteAllText((Join-Path $integration $file),'<Project/>') }
        $oldCL = $env:CL; $oldLINK = $env:LINK; $oldPath = $env:Path
        try {
            $env:CL = '/DUNREVIEWED_AMBIENT /FIambient.h'
            $env:LINK = '/FORCE:UNRESOLVED /LIBPATH:ambient'
            $environment = Get-RSVcpkgBuildEnvironment -RepoRoot $root -ToolRoot $tool -Platform x64 -Configuration Release
            $child = New-RSProcessStartInfo -FilePath 'fixture.exe' -WorkingDirectory $root -AdditionalEnvironment ([hashtable]$environment)
            $child.Environment['CL'] | Should Be ('/I"{0}"' -f (Join-Path $installed 'include'))
            $child.Environment['LINK'] | Should Match 'manual-link'
            $child.Environment['LINK'] | Should Not Match 'UNRESOLVED|ambient'
            $child.Environment['VcpkgManifestInstall'] | Should Be 'false'
            $child.Environment['ForceImportBeforeCppTargets'] | Should Be (Join-Path $integration 'vcpkg.props')
            $env:CL | Should Be '/DUNREVIEWED_AMBIENT /FIambient.h'
            $env:LINK | Should Be '/FORCE:UNRESOLVED /LIBPATH:ambient'
            $env:Path | Should Be $oldPath
            Remove-Item -LiteralPath (Join-Path $integration 'vcpkg.targets')
            [void][IO.Directory]::CreateDirectory((Join-Path $integration 'vcpkg.targets'))
            Get-BootstrapFailure { Get-RSVcpkgBuildEnvironment -RepoRoot $root -ToolRoot $tool -Platform x64 -Configuration Release } | Should Match 'Initialize-RSVcpkgToolContext.*vcpkg-install.ps1 -Platform x64'
        } finally {
            $env:CL = $oldCL; $env:LINK = $oldLINK
        }
    }

    It 'provisions dependencies only for native build roles before the first build invocation' {
        $adapter = [IO.File]::ReadAllText((Join-Path $repoRoot 'Tools/Invoke-ProjectCi.ps1'))
        $staticEnd = $adapter.IndexOf("if ($" + "Phase -ne 'Qualification')")
        $provision = $adapter.IndexOf('Initialize-RSVcpkgToolContext')
        $install = $adapter.IndexOf("Run 'vcpkg-install.ps1'")
        $build = $adapter.IndexOf("Run 'build.ps1'")
        $provision | Should BeGreaterThan $staticEnd
        $install | Should BeGreaterThan $provision
        $build | Should BeGreaterThan $install
        $adapter | Should Match "'ReleaseTestBuild','Production','Qualification'"
        foreach ($path in @('.github/workflows/build-reusable.yml','.github/workflows/ghostty-runtime-product-qualification.yml')) {
            [IO.File]::ReadAllText((Join-Path $repoRoot $path)) | Should Match 'Initialize-RSVcpkgToolContext'
        }
    }
}
