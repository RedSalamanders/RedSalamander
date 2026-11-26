Import-Module (Join-Path $PSScriptRoot 'TestSupport.psm1') -Force -ErrorAction Stop
$repoRoot=Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
Import-Module (Join-Path $repoRoot 'Tools/Modules/Testing/ScopedTesting.psm1') -Force
Import-Module (Join-Path $repoRoot 'Tools/Modules/Testing/TestSandbox.psm1') -Force
Describe 'Scoped testing entry points' {
    It 'validates real fragments before affected execution and keeps explanation non-executing' {
        $fixture=New-RSTestSandboxScratchDirectory -RepoRoot $repoRoot -Harness 'tools-pester' -Case ('fragment-admission-'+[guid]::NewGuid().ToString('N'))
        $moduleDirectory=Join-Path $fixture 'Tools/Modules/Testing'
        [void](New-Item -ItemType Directory -Path $moduleDirectory,(Join-Path $fixture '.ci'),(Join-Path $fixture 'Changes') -Force)
        foreach($relative in @('Test-Changes.ps1','test.ps1','Tools/Test-Changelog.ps1','.ci/project.json')) {
            Copy-Item -LiteralPath (Join-Path $repoRoot $relative) -Destination (Join-Path $fixture $relative)
        }
        # The Tools/Test-Changelog.ps1 engine shim runs the pinned engine from the fixture's own .ci-engine.
        Invoke-RSTestFixtureGit -Arguments @('clone';'--quiet';'--no-hardlinks';(Join-Path $repoRoot '.ci-engine');(Join-Path $fixture '.ci-engine')) *> (Join-Path $fixture 'engine-clone.log')
        $LASTEXITCODE | Should Be 0
        'function Assert-ScopedTestNames { param($Root) return 1 }; Export-ModuleMember -Function Assert-ScopedTestNames' |
            Set-Content -LiteralPath (Join-Path $moduleDirectory 'ScopedTesting.psm1')
        @'
param([switch]$ExplainPlan)
'selected' | Set-Content -LiteralPath (Join-Path $PSScriptRoot '../launched.txt')
$global:LASTEXITCODE=2
'@ | Set-Content -LiteralPath (Join-Path $fixture 'Tools/Run-AllTests.ps1')
        'Fragment instructions' | Set-Content -LiteralPath (Join-Path $fixture 'Changes/README.md')
        "# Changelog`n`n## Unreleased`n" | Set-Content -LiteralPath (Join-Path $fixture 'CHANGELOG.md')
        $fragment=Join-Path $fixture 'Changes/2026-11-09-example.md'
        '- Valid fixture change.' | Set-Content -LiteralPath $fragment
        $child=Join-Path $PSHOME 'pwsh.exe'
        Invoke-RSTestFixtureNative -FilePath $child -Arguments @('-NoProfile'; '-File'; (Join-Path $fixture 'Test-Changes.ps1')) *> (Join-Path $fixture 'valid.log')
        $LASTEXITCODE | Should Be 2
        $launched=Join-Path $fixture 'launched.txt'
        (Test-Path -LiteralPath $launched) | Should Be $true
        Remove-Item -LiteralPath $launched
        'Malformed fixture with no bullet.' | Set-Content -LiteralPath $fragment
        Invoke-RSTestFixtureNative -FilePath $child -Arguments @('-NoProfile'; '-File'; (Join-Path $fixture 'Test-Changes.ps1')) *> (Join-Path $fixture 'invalid.log')
        ($LASTEXITCODE -ne 0) | Should Be $true
        (Test-Path -LiteralPath $launched) | Should Be $false
        Invoke-RSTestFixtureNative -FilePath $child -Arguments @('-NoProfile'; '-File'; (Join-Path $fixture 'Test-Changes.ps1'); '-Explain') *> (Join-Path $fixture 'explain.log')
        $LASTEXITCODE | Should Be 2
        (Test-Path -LiteralPath $launched) | Should Be $true
        $global:LASTEXITCODE=0
    }
    It 'preserves legacy selection and forwards the root defaults, Full, and alternate-volume opt-in' {
        $fixture=New-RSTestSandboxScratchDirectory -RepoRoot $repoRoot -Harness 'tools-pester' -Case ('alternate-forwarding-'+[guid]::NewGuid().ToString('N'))
        $succeeded=$false
        try {
            $moduleDirectory=Join-Path $fixture 'Tools/Modules/Testing'
            [void](New-Item -ItemType Directory -Path $moduleDirectory -Force)
            foreach ($entry in 'Test-Changes.ps1','test.ps1') {
                Copy-Item -LiteralPath (Join-Path $repoRoot $entry) -Destination (Join-Path $fixture $entry)
            }
            'param($Base)' | Set-Content -LiteralPath (Join-Path $fixture 'Tools/Test-Changelog.ps1')
            @'
function Assert-ScopedTestNames { param($Root) return 1 }
Export-ModuleMember -Function Assert-ScopedTestNames
'@ | Set-Content -LiteralPath (Join-Path $moduleDirectory 'ScopedTesting.psm1')
            @'
param($Suite,$Configuration,$Platform,$ImpactBase,[switch]$SkipBuild,[switch]$NonInteractive,$ValidationMode,$TestRoot,
      [switch]$AllowAlternateVolumeTestRoot,$CaseFilter='',$CommandsFamily='',[int]$BuildNumber=0)
[ordered]@{
    suite=$Suite; mode=$ValidationMode; testRoot=$TestRoot; caseFilter=$CaseFilter
    alternate=[bool]$AllowAlternateVolumeTestRoot; nonInteractive=[bool]$NonInteractive; buildNumber=$BuildNumber
} |
    ConvertTo-Json | Set-Content -LiteralPath (Join-Path $PSScriptRoot '../captured.json')
$global:LASTEXITCODE=2
'@ | Set-Content -LiteralPath (Join-Path $fixture 'Tools/Run-AllTests.ps1')
            $cases = @(
                @{ name='LegacyFull'; script='Test-Changes.ps1'; arguments=@('-Mode','Full'); suite='Full'; mode='Fresh'; filter=''; number=0 }
                @{ name='LegacyPrePush'; script='Test-Changes.ps1'; arguments=@('-Mode','PrePush'); suite='Full'; mode='Fresh'; filter=''; number=0 }
                @{ name='LegacyFocused'; script='Test-Changes.ps1'; arguments=@('-Suite','FileOps','-CaseFilter','FileOps_IndependentVolumes'); suite='FileOps'; mode='Fresh'; filter='FileOps_IndependentVolumes'; number=0 }
                @{ name='RootAffected'; script='test.ps1'; arguments=@(); suite='Full'; mode='Affected'; filter=''; number=0 }
                @{ name='RootFull'; script='test.ps1'; arguments=@('-Full','-BuildNumber','42'); suite='Full'; mode='Fresh'; filter=''; number=42 }
            )
            foreach ($case in $cases) {
                foreach ($optIn in @($false,$true)) {
                    $arguments = @('-NoProfile','-File',(Join-Path $fixture $case.script),'-TestRoot',$env:REDSALAMANDER_TEST_ROOT) + $case.arguments
                    if ($optIn) { $arguments += '-AllowAlternateVolumeTestRoot' }
                    Invoke-RSTestFixtureNative -FilePath (Join-Path $PSHOME 'pwsh.exe') -Arguments $arguments *> (Join-Path $fixture "child-$($case.name)-$optIn.log")
                    $LASTEXITCODE | Should Be 2
                    $captured = Get-Content -LiteralPath (Join-Path $fixture 'captured.json') -Raw | ConvertFrom-Json
                    $captured.alternate | Should Be $optIn
                    $captured.suite | Should Be $case.suite
                    $captured.mode | Should Be $case.mode
                    $captured.caseFilter | Should Be $case.filter
                    $captured.buildNumber | Should Be $case.number
                    $captured.nonInteractive | Should Be $true
                    $captured.testRoot | Should Be $env:REDSALAMANDER_TEST_ROOT
                }
            }
            $succeeded=$true
        } finally {
            if($succeeded) {
                [void](Assert-RSTestSandboxContainedPath -TestRoot $env:REDSALAMANDER_TEST_ROOT -Path $fixture -RequireDescendant)
                $refusals=@(& (Get-Module TestSandbox) { param($Path) Remove-RSTestSandboxTreeWithoutFollowingLinks -Path $Path } $fixture)
                if($refusals.Count){throw 'Scoped-wrapper fixture cleanup was refused.'}
            }
            $global:LASTEXITCODE=0
        }
    }
    It 'accounts for all active native sources with contextual test names' {(Assert-ScopedTestNames $repoRoot) | Should BeGreaterThan 0}
}

Describe 'Canonical scoped evidence authority' {
    It 'keeps obsolete receipt and source identity exports absent' {
        $exports = @((Get-Module ScopedTesting).ExportedCommands.Keys)
        foreach ($name in @('Get-ScopedSourceIdentity','Get-ScopedArtifactIdentity','Get-ScopedRunIdentity',
                'Get-ScopedEnvironmentIdentity','Test-ScopedReceipt','Write-ScopedReceipt')) {
            ($exports -contains $name) | Should Be $false
        }
        $runner = [IO.File]::ReadAllText((Join-Path $repoRoot 'Tools/Run-AllTests.ps1'))
        $runner | Should Match 'hosted_stock_preparation=\[string\]\$env:RS_HOSTED_ARM64_STOCK_PREPARATION'
    }
}
Describe 'Engine-free scoped metadata transport' {
    It 'preserves raw NUL Unicode and newline payloads without a CI engine or caller mutation' {
        $fixture=New-RSTestSandboxScratchDirectory -RepoRoot $repoRoot -Harness 'tools-pester' -Case ('scoped-native-'+[guid]::NewGuid().ToString('N'))
        $copy=Join-Path $fixture 'ScopedBootstrap.psm1'
        Copy-Item -LiteralPath (Join-Path $repoRoot 'Tools/Modules/Testing/ScopedTesting.psm1') -Destination $copy
        (Test-Path (Join-Path $fixture '.ci-engine'))|Should Be $false
        $module=Import-Module $copy -PassThru -Scope Local -Prefix ScopedBootstrap
        $cwd=(Get-Location).Path;$prior=[Environment]::GetEnvironmentVariable('RS_METADATA_PARENT_SENTINEL','Process')
        $savedConsoleEncoding=[Console]::OutputEncoding
        try {
            [Console]::OutputEncoding=[Text.UTF8Encoding]::new($false)
            [Environment]::SetEnvironmentVariable('RS_METADATA_PARENT_SENTINEL','fixture é雪','Process')
            $value='raw é雪'+[char]13+[char]10+'middle'+[char]10
            $actual=& $module {param($root,$value)Invoke-ScopedGit $root @('-c',('fixture.payload='+$value),'config','--null','--get','fixture.payload')} $fixture $value
            $actual|Should Be ($value+[char]0)
            (Get-Location).Path|Should Be $cwd
            [Environment]::GetEnvironmentVariable('RS_METADATA_PARENT_SENTINEL','Process')|Should Be 'fixture é雪'
            [Console]::OutputEncoding=[Text.Encoding]::GetEncoding(437)
            $legacy=& $module {param($root,$value)Invoke-ScopedGit $root @('-c',('fixture.payload='+$value),'config','--null','--get','fixture.payload')} $fixture $value
            $legacy|Should Be ($value+[char]0)
            [Console]::OutputEncoding=[Text.UTF8Encoding]::new($false)
            $null=& $module {param($root)Invoke-ScopedGit $root @('init','--quiet')} $fixture
            $path='café-Δ-雪.txt'
            [IO.File]::WriteAllText((Join-Path $fixture $path),'owned fixture',[Text.UTF8Encoding]::new($false))
            foreach($encoding in @([Text.Encoding]::GetEncoding(437),[Text.UTF8Encoding]::new($false))) {
                [Console]::OutputEncoding=$encoding
                $paths=@(& $module {param($root)Get-ScopedTrackedPaths $root} $fixture)
                ($paths -ccontains $path)|Should Be $true
                [Console]::OutputEncoding.CodePage|Should Be $encoding.CodePage
            }
            $failed=$false
            try {& $module {param($root)Invoke-ScopedGit $root @('rev-parse','--verify','missing-object')} $fixture}
            catch {$failed=$_.Exception.Message.StartsWith('Git scope discovery failed: ')}
            $failed|Should Be $true
        } finally {
            [Environment]::SetEnvironmentVariable('RS_METADATA_PARENT_SENTINEL',$prior,'Process')
            [Console]::OutputEncoding=$savedConsoleEncoding
            Remove-Module $module
        }
    }
}
