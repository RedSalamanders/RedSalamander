Set-StrictMode -Version Latest

function Assert-RedMonitorClientPath([string] $Path) {
    $current = [IO.Path]::GetPathRoot($Path)
    foreach ($part in $Path.Substring($current.Length).Split([char[]]@('\','/'))) {
        $current = Join-Path $current $part
        try { $item = Get-Item -LiteralPath $current -Force -ErrorAction Stop }
        catch [Management.Automation.ItemNotFoundException] { return }
        if ($item.Attributes -band [IO.FileAttributes]::ReparsePoint) {
            throw "Reparse RedMonitor client input: $current"
        }
    }
}

function Get-RedMonitorClientCommand {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string] $Root,
        [Parameter(Mandatory)][ValidateSet('Update-RedMonitorClient.ps1','Verify-RedMonitorClient.ps1')][string] $Name
    )

    $rootPath = [IO.Path]::GetFullPath($Root).TrimEnd('\','/')
    $lock = Join-Path $rootPath 'Dependencies/RedMonitor.lock.json'
    $directory = Join-Path $rootPath 'Dependencies/RedMonitorClientTools'
    $names = @('Update-RedMonitorClient.ps1','Verify-RedMonitorClient.ps1')
    $paths = @($lock) + @($names | ForEach-Object { Join-Path $directory $_ })
    $vendor = Join-Path $rootPath 'ThirdParty/RedMonitor'
    $paths += @('include/RedMonitor/Client.h','LICENSE.txt','.gitattributes' | ForEach-Object { Join-Path $vendor $_ })
    foreach ($path in $paths) { Assert-RedMonitorClientPath $path }

    $utf8 = [Text.UTF8Encoding]::new($false,$true)
    $pin = ConvertFrom-Json -InputObject ([IO.File]::ReadAllText($lock,$utf8)) -ErrorAction Stop
    if (-not $pin -or $pin.repository -cne 'https://github.com/RedSalamanders/RedMonitor' -or
        -not $pin.PSObject.Properties['tooling']) {
        throw "Missing immutable RedMonitor client tooling pin in $lock"
    }
    $tooling = $pin.tooling
    if (-not $tooling -or $tooling.schemaVersion -ne 1 -or $tooling.commit -cnotmatch '^[a-f0-9]{40}$' -or
        @($tooling.files).Count -ne $names.Count) {
        throw "Invalid immutable RedMonitor client tooling pin in $lock"
    }
    $selected = $null
    $sha = [Security.Cryptography.SHA256]::Create()
    try {
        foreach ($command in $names) {
            $entries = @($tooling.files | Where-Object { $_.path -ceq $command })
            if ($entries.Count -ne 1 -or $entries[0].sha256 -cnotmatch '^[a-f0-9]{64}$') {
                throw "Invalid RedMonitor client tooling entry: $command"
            }
            $file = Join-Path $directory $command
            if (-not (Test-Path -LiteralPath $file -PathType Leaf)) {
                throw "Missing pinned RedMonitor client tool: $command"
            }
            $bytes = [IO.File]::ReadAllBytes($file)
            $actual = [BitConverter]::ToString($sha.ComputeHash($bytes)).Replace('-', '')
            if ($actual -ine $entries[0].sha256) { throw "RedMonitor client tooling hash mismatch: $command" }
            if ($command -ceq $Name) { $selected = $bytes }
        }
    } finally { $sha.Dispose() }

    return [scriptblock]::Create($utf8.GetString($selected))
}
Export-ModuleMember -Function Get-RedMonitorClientCommand
