Describe 'Release compiler and final-link policy' {
    It 'isolates ordinary native profiles from the ambient Embedded recovery format and rejects broken controls' {
        $hadFormat = Test-Path Env:RSBuildCompilerDebugInformation
        $originalFormat = $env:RSBuildCompilerDebugInformation
        try {
            $env:RSBuildCompilerDebugInformation = 'Embedded'
            & (Join-Path $PSScriptRoot 'ReleaseBuildPolicyCheck.ps1')
            $env:RSBuildCompilerDebugInformation | Should Be 'Embedded'
        } finally {
            if ($hadFormat) {
                $env:RSBuildCompilerDebugInformation = $originalFormat
            } else {
                Remove-Item Env:RSBuildCompilerDebugInformation -ErrorAction SilentlyContinue
            }
        }
    }
}
