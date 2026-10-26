Describe 'Release compiler and final-link policy' {
    It 'evaluates native profiles and rejects disabled compiler and linker controls' {
        & (Join-Path $PSScriptRoot 'ReleaseBuildPolicyCheck.ps1')
    }
}