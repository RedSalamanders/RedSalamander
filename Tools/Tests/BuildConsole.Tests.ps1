Describe 'Native build console and short dependency paths' {
    It 'refuses warnings and collisions, emits safe links and never auto-runs output' {
        & (Join-Path $PSScriptRoot 'BuildConsoleCheck.ps1')
    }
}