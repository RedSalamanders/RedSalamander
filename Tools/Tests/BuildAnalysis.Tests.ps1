$repoRoot = (Resolve-Path (Join-Path $PSScriptRoot '../..')).Path
$analysisGuard = Join-Path $repoRoot 'Tools/Assert-AnalysisOutput.ps1'
$rootBuild = Join-Path $repoRoot 'build.ps1'

Describe 'Isolated native code analysis' {
    BeforeEach {
        $runId = [Guid]::NewGuid().ToString('N')
        $runRoot = Join-Path $repoRoot ".build/analysis/$runId"
        $guardArgs = @{
            RunId = $runId
            OutputRoot = $runRoot
            OutputDirectory = Join-Path $runRoot 'x64/Debug'
            IntermediateDirectory = Join-Path $runRoot 'Intermediate/x64/Debug/Common'
        }
    }

    It 'accepts owned diagnostic paths without creating outputs' {
        & $analysisGuard @guardArgs
        Test-Path -LiteralPath $runRoot | Should Be $false
    }

    It 'refuses the ordinary native output root' {
        $guardArgs.OutputRoot = Join-Path $repoRoot '.build'
        { & $analysisGuard @guardArgs } | Should Throw 'guarded repository root'
    }

    It 'refuses a path from another analysis run' {
        $otherRun = Join-Path $repoRoot ".build/analysis/$([Guid]::NewGuid().ToString('N'))"
        $guardArgs.IntermediateDirectory = Join-Path $otherRun 'Intermediate/x64/Debug/Common'
        { & $analysisGuard @guardArgs } | Should Throw 'guarded repository root'
    }

    It 'rejects output inside the intermediate tree' {
        $guardArgs.OutputDirectory = $guardArgs.IntermediateDirectory
        { & $analysisGuard @guardArgs } | Should Throw 'distinct descendants'
    }

    It 'rejects Clean and production flags before a compiler can be selected' {
        $previous = $env:CI_NATIVE_BUILDS_DISABLED
        try {
            $env:CI_NATIVE_BUILDS_DISABLED = 'true'
            { & $rootBuild -Analyze -Clean } | Should Throw '-Analyze requires a source build'
            { & $rootBuild -Analyze -Rebuild } | Should Throw '-Analyze requires a source build'
            { & $rootBuild -Analyze -Zip } | Should Throw '-Analyze requires a source build'
            { & $rootBuild -Analyze -ProductionOnly } | Should Throw '-Analyze requires a source build'
            { & $rootBuild -Analyze -DxUiSdkInputDescriptor 'missing.json' } | Should Throw '-Analyze requires a source build'
            { & $rootBuild -Analyze } | Should Throw 'Native builds are forbidden'
        }
        finally {
            $env:CI_NATIVE_BUILDS_DISABLED = $previous
        }
    }
}
