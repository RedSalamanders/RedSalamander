Set-StrictMode -Version Latest

function Resolve-RSRedMonitorTestInput {
    <# Historical caller compatibility. Only consumer-local source is read; viewer contracts live upstream. #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$RepoRoot,[Parameter(Mandatory)][string]$Path)
    $root=[IO.Path]::GetFullPath($RepoRoot)
    $full=[IO.Path]::GetFullPath($Path,$root)
    $relative=[IO.Path]::GetRelativePath($root,$full).Replace('\','/')
    if($relative -match '^(RedSalamanderMonitor/|Tests/MonitorTest/)'){
        throw 'Viewer implementation/tests belong to the independent RedMonitor repository, not consumer validation.'
    }
    return $full
}
Export-ModuleMember -Function Resolve-RSRedMonitorTestInput