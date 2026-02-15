# Initialize ETW permissions for RedSalamanderMonitor
# Adds (or removes) a user from the local "Performance Log Users" group so ETW sessions can start without elevation.

param(
    [string]$User = "$env:USERDOMAIN\$env:USERNAME",
    [switch]$Remove,
    [switch]$NoPause
)

function Test-IsAdministrator {
    $identity = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object Security.Principal.WindowsPrincipal($identity)
    return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

function Get-PerformanceLogUsersGroupName {
    # Built-in group SID: S-1-5-32-559 (Performance Log Users). Translate to localized name.
    $sid = New-Object System.Security.Principal.SecurityIdentifier("S-1-5-32-559")
    $account = $sid.Translate([System.Security.Principal.NTAccount]).Value

    $slashIndex = $account.LastIndexOf('\')
    if ($slashIndex -ge 0 -and $slashIndex -lt ($account.Length - 1)) {
        return $account.Substring($slashIndex + 1)
    }

    return $account
}

function Relaunch-Elevated {
    $args = @("-NoProfile", "-ExecutionPolicy", "Bypass", "-File", "`"$PSCommandPath`"")
    if ($User) { $args += @("-User", $User) }
    if ($Remove) { $args += "-Remove" }
    if ($NoPause) { $args += "-NoPause" }

    Start-Process -FilePath "powershell.exe" -Verb RunAs -ArgumentList $args | Out-Null
    exit 0
}

Write-Host "=== ETW Permission Setup (RedSalamanderMonitor) ===" -ForegroundColor Cyan

$groupName = Get-PerformanceLogUsersGroupName
Write-Host "User:  $User" -ForegroundColor Gray
Write-Host "Group: $groupName" -ForegroundColor Gray
Write-Host ""

if (-not (Test-IsAdministrator)) {
    Write-Host "Elevation required to modify local group membership." -ForegroundColor Yellow
    Write-Host "Requesting Administrator privileges..." -ForegroundColor Yellow
    Write-Host ""
    Relaunch-Elevated
}

if ($Remove) {
    Write-Host "Removing user from '$groupName'..." -ForegroundColor Gray
    & net localgroup "$groupName" "$User" /delete | Out-Null
    if ($LASTEXITCODE -eq 0) {
        Write-Host "OK: Removed successfully." -ForegroundColor Green
    } else {
        Write-Host "ERROR: Failed to remove user (exit code: $LASTEXITCODE)." -ForegroundColor Red
        Write-Host "Tip: Verify the user name and that it is currently a member of the group." -ForegroundColor Yellow
    }
} else {
    Write-Host "Adding user to '$groupName'..." -ForegroundColor Gray
    & net localgroup "$groupName" "$User" /add | Out-Null
    if ($LASTEXITCODE -eq 0) {
        Write-Host "OK: Added successfully." -ForegroundColor Green
    } else {
        Write-Host "ERROR: Failed to add user (exit code: $LASTEXITCODE)." -ForegroundColor Red
        Write-Host "Tip: Verify the user name and that this machine supports local group management." -ForegroundColor Yellow
    }
}

Write-Host ""
Write-Host "IMPORTANT: Sign out/in (or reboot) so your access token picks up the new group membership." -ForegroundColor Cyan
Write-Host "Then launch RedSalamanderMonitor normally (no UAC prompt expected on most machines)." -ForegroundColor Cyan

Write-Host ""
if (-not $NoPause) {
    Write-Host "Press any key to continue..."
    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
}
