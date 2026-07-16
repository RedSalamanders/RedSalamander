#Requires -Version 5.1
<#
.SYNOPSIS
    Lists Windows Credential Manager entries and optionally removes them by regex.

.DESCRIPTION
    Testing accumulates many generic credentials in the Windows Credential
    Manager (the same store RedSalamander writes to via CredWriteW with
    CRED_TYPE_GENERIC). This tool enumerates those credentials through the
    advapi32 CredEnumerate/CredDelete APIs and can prune them by a regular
    expression matched against the credential TargetName.

    LISTING IS THE DEFAULT and is always read-only. Removal only happens when
    -Remove is supplied, and it honors -WhatIf / -Confirm.

    Two safety nets scope the default selection to RedSalamander test debris:
      * -Pattern defaults to '^RedSalamander/', the target-name namespace the
        app uses (see ConnectionSecrets.cpp BuildCredentialTargetName).
      * Only Generic credentials are considered unless -IncludeAllTypes is set,
        so domain / Windows / web credentials are never touched by accident.

    Use -All to clear the RedSalamander default and inspect every credential,
    and -Pattern to supply your own regex.

.PARAMETER Pattern
    Regular expression matched (case-insensitively by default) against each
    credential's TargetName. Defaults to '^RedSalamander/'. Combine with -All
    to widen, or pass '.' to match everything.

.PARAMETER All
    Clear the default RedSalamander namespace filter. With -All and no -Pattern,
    every credential is matched. With -All and an explicit -Pattern, only your
    pattern applies.

.PARAMETER Remove
    Delete the matched credentials instead of just listing them. Honors -WhatIf
    and -Confirm; you are prompted before each deletion unless you answer
    "Yes to All".

.PARAMETER IncludeAllTypes
    Consider every credential type (DomainPassword, DomainCertificate, etc.),
    not just Generic. Off by default so non-test credentials stay protected.

.PARAMETER CaseSensitive
    Match -Pattern case-sensitively.

.PARAMETER PassThru
    Emit the matched credentials as objects to the pipeline in addition to the
    on-screen table, so they can be filtered or counted by the caller.

.EXAMPLE
    .\Tools\Manage-WindowsCredentials.ps1
    Lists every generic credential whose target starts with "RedSalamander/".

.EXAMPLE
    .\Tools\Manage-WindowsCredentials.ps1 -All
    Lists every generic credential in the current user's store.

.EXAMPLE
    .\Tools\Manage-WindowsCredentials.ps1 -Remove -WhatIf
    Shows which RedSalamander test credentials *would* be removed.

.EXAMPLE
    .\Tools\Manage-WindowsCredentials.ps1 -Pattern '^RedSalamander/Connections/.*/password$' -Remove
    Removes only the stored connection passwords, after confirmation.

.EXAMPLE
    .\Tools\Manage-WindowsCredentials.ps1 -All -Pattern 'sftp-test' -Remove -Confirm:$false
    Removes every credential whose target contains "sftp-test", no prompt.
#>

[CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'High')]
param(
    [Parameter(Position = 0)]
    [string] $Pattern,

    [switch] $All,

    [switch] $Remove,

    [switch] $IncludeAllTypes,

    [switch] $CaseSensitive,

    [switch] $PassThru,

    [Alias('h', '?')]
    [switch] $Help
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Show-Usage {
    Write-Host 'Manage-WindowsCredentials.ps1' -ForegroundColor White
    Write-Host 'List Windows Credential Manager entries and optionally remove them by regex.' -ForegroundColor Cyan
    Write-Host ''
    Write-Host 'USAGE' -ForegroundColor White
    Write-Host '  .\Tools\Manage-WindowsCredentials.ps1 [-Pattern <regex>] [-All] [-IncludeAllTypes]'
    Write-Host '                                        [-CaseSensitive] [-PassThru]'
    Write-Host '  .\Tools\Manage-WindowsCredentials.ps1 [-Pattern <regex>] -Remove [-WhatIf] [-Confirm:$false]'
    Write-Host '  .\Tools\Manage-WindowsCredentials.ps1 -Help'
    Write-Host ''
    Write-Host 'PARAMETERS' -ForegroundColor White
    Write-Host '  -Pattern <regex>  Regex matched against each credential TargetName.'
    Write-Host "                    Default: '^RedSalamander/' (the app's test namespace)."
    Write-Host '  -All              Drop the default RedSalamander filter; with no -Pattern, match everything.'
    Write-Host '  -Remove           Delete the matched credentials instead of just listing them.'
    Write-Host '  -WhatIf           Preview deletions without applying (use with -Remove).'
    Write-Host '  -Confirm:$false   Skip the per-credential confirmation prompt (use with -Remove).'
    Write-Host '  -IncludeAllTypes  Consider all credential types, not just Generic (off by default).'
    Write-Host '  -CaseSensitive    Match -Pattern case-sensitively.'
    Write-Host '  -PassThru         Also emit the matched credentials as objects to the pipeline.'
    Write-Host '  -Help, -h, -?     Show this help.'
    Write-Host ''
    Write-Host 'SAFETY' -ForegroundColor White
    Write-Host '  Listing is the default and is always read-only. Removal requires -Remove and'
    Write-Host '  honors -WhatIf / -Confirm. By default only Generic credentials whose target'
    Write-Host '  starts with "RedSalamander/" are in scope, so domain / Windows / web logins'
    Write-Host '  are never touched unless you widen the selection with -All / -IncludeAllTypes.'
    Write-Host ''
    Write-Host 'EXAMPLES' -ForegroundColor White
    Write-Host '  .\Tools\Manage-WindowsCredentials.ps1'
    Write-Host '      List RedSalamander test credentials (read-only).'
    Write-Host '  .\Tools\Manage-WindowsCredentials.ps1 -All'
    Write-Host '      List every generic credential in the store.'
    Write-Host '  .\Tools\Manage-WindowsCredentials.ps1 -Remove -WhatIf'
    Write-Host '      Preview which RedSalamander test credentials would be removed.'
    Write-Host '  .\Tools\Manage-WindowsCredentials.ps1 -Remove'
    Write-Host '      Remove the RedSalamander test credentials, prompting per item.'
    Write-Host "  .\Tools\Manage-WindowsCredentials.ps1 -Pattern '^RedSalamander/Connections/.*/password$' -Remove"
    Write-Host '      Remove only the stored connection passwords.'
    Write-Host "  .\Tools\Manage-WindowsCredentials.ps1 -All -Pattern 'sftp-test' -Remove -Confirm:`$false"
    Write-Host '      Remove every credential whose target contains "sftp-test", no prompt.'
}

if ($Help) {
    Show-Usage
    return
}

# ---------------------------------------------------------------------------
# Native interop: enumerate and delete Windows Credential Manager entries.
# ---------------------------------------------------------------------------
if (-not ([System.Management.Automation.PSTypeName]'RedSalamander.CredentialManager').Type) {
    Add-Type -Namespace 'RedSalamander' -Name 'CredentialManager' -UsingNamespace 'System.Collections.Generic' -MemberDefinition @'
[StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
private struct CREDENTIAL
{
    public uint Flags;
    public uint Type;
    public IntPtr TargetName;
    public IntPtr Comment;
    public System.Runtime.InteropServices.ComTypes.FILETIME LastWritten;
    public uint CredentialBlobSize;
    public IntPtr CredentialBlob;
    public uint Persist;
    public uint AttributeCount;
    public IntPtr Attributes;
    public IntPtr TargetAlias;
    public IntPtr UserName;
}

[DllImport("advapi32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
private static extern bool CredEnumerate(string filter, uint flags, out uint count, out IntPtr pCredentials);

[DllImport("advapi32.dll", SetLastError = true)]
private static extern void CredFree(IntPtr buffer);

[DllImport("advapi32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
private static extern bool CredDelete(string target, uint type, uint flags);

private const int ERROR_NOT_FOUND = 1168;

public sealed class CredInfo
{
    public string TargetName;
    public string UserName;
    public uint Type;
    public uint Persist;
    public uint BlobSize;
    public DateTime LastWritten;
}

private static DateTime FileTimeToDateTime(System.Runtime.InteropServices.ComTypes.FILETIME ft)
{
    long value = ((long)(uint)ft.dwHighDateTime << 32) | (uint)ft.dwLowDateTime;
    if (value <= 0) { return DateTime.MinValue; }
    try { return DateTime.FromFileTimeUtc(value).ToLocalTime(); }
    catch { return DateTime.MinValue; }
}

public static List<CredInfo> Enumerate()
{
    uint count;
    IntPtr pCredentials;
    if (!CredEnumerate(null, 0, out count, out pCredentials))
    {
        int err = Marshal.GetLastWin32Error();
        if (err == ERROR_NOT_FOUND) { return new List<CredInfo>(); }
        throw new System.ComponentModel.Win32Exception(err);
    }

    try
    {
        var result = new List<CredInfo>((int)count);
        for (int i = 0; i < count; i++)
        {
            IntPtr pCred = Marshal.ReadIntPtr(pCredentials, i * IntPtr.Size);
            CREDENTIAL cred = (CREDENTIAL)Marshal.PtrToStructure(pCred, typeof(CREDENTIAL));
            result.Add(new CredInfo
            {
                TargetName  = cred.TargetName != IntPtr.Zero ? Marshal.PtrToStringUni(cred.TargetName) : null,
                UserName    = cred.UserName   != IntPtr.Zero ? Marshal.PtrToStringUni(cred.UserName)   : null,
                Type        = cred.Type,
                Persist     = cred.Persist,
                BlobSize    = cred.CredentialBlobSize,
                LastWritten = FileTimeToDateTime(cred.LastWritten)
            });
        }
        return result;
    }
    finally
    {
        CredFree(pCredentials);
    }
}

public static void Delete(string targetName, uint type)
{
    if (!CredDelete(targetName, type, 0))
    {
        throw new System.ComponentModel.Win32Exception(Marshal.GetLastWin32Error());
    }
}
'@
}

function Get-CredTypeName([uint32] $Type) {
    switch ($Type) {
        1 { 'Generic' }
        2 { 'DomainPassword' }
        3 { 'DomainCertificate' }
        4 { 'DomainVisiblePassword' }
        5 { 'GenericCertificate' }
        6 { 'DomainExtended' }
        default { "Type$Type" }
    }
}

function Get-CredPersistName([uint32] $Persist) {
    switch ($Persist) {
        1 { 'Session' }
        2 { 'LocalMachine' }
        3 { 'Enterprise' }
        default { "Persist$Persist" }
    }
}

# ---------------------------------------------------------------------------
# Selection.
# ---------------------------------------------------------------------------
$effectivePattern = if ($PSBoundParameters.ContainsKey('Pattern') -and $Pattern) {
    $Pattern
} elseif ($All) {
    '.'                       # -All with no explicit pattern: match every target.
} else {
    '^RedSalamander/'         # Safe default: only RedSalamander test credentials.
}

$regexOptions = if ($CaseSensitive) {
    [System.Text.RegularExpressions.RegexOptions]::None
} else {
    [System.Text.RegularExpressions.RegexOptions]::IgnoreCase
}

try {
    $regex = [System.Text.RegularExpressions.Regex]::new($effectivePattern, $regexOptions)
} catch {
    throw "Invalid -Pattern regular expression '$effectivePattern': $($_.Exception.Message)"
}

$allCreds = [RedSalamander.CredentialManager]::Enumerate()

$matched = @(
    $allCreds | Where-Object {
        $_.TargetName -and
        ($IncludeAllTypes -or $_.Type -eq 1) -and
        $regex.IsMatch($_.TargetName)
    } | Sort-Object TargetName
)

$scope = if ($IncludeAllTypes) { 'all types' } else { 'Generic only' }
Write-Host ("Store: {0} credential(s) total. Pattern '{1}' ({2}, {3}) matched {4}." -f `
        $allCreds.Count, $effectivePattern, $(if ($CaseSensitive) { 'case-sensitive' } else { 'case-insensitive' }), $scope, $matched.Count) -ForegroundColor Cyan

if ($matched.Count -eq 0) {
    Write-Host 'Nothing to do.' -ForegroundColor DarkGray
    return
}

# ---------------------------------------------------------------------------
# List (always) — read-only table.
# ---------------------------------------------------------------------------
$rows = $matched | ForEach-Object {
    [pscustomobject]@{
        TargetName  = $_.TargetName
        UserName    = if ($_.UserName) { $_.UserName } else { '(none)' }
        Type        = Get-CredTypeName $_.Type
        Persist     = Get-CredPersistName $_.Persist
        Bytes       = $_.BlobSize
        LastWritten = if ($_.LastWritten -eq [datetime]::MinValue) { '' } else { $_.LastWritten.ToString('yyyy-MM-dd HH:mm:ss') }
    }
}

$rows | Format-Table -AutoSize | Out-Host

if ($PassThru) {
    $matched
}

# ---------------------------------------------------------------------------
# Remove (opt-in) — honors -WhatIf / -Confirm via ShouldProcess.
# ---------------------------------------------------------------------------
if (-not $Remove) {
    Write-Host ''
    Write-Host "Listing only. Re-run with -Remove to delete these $($matched.Count) credential(s) (add -WhatIf to preview)." -ForegroundColor Yellow
    return
}

$removed = 0
$failed = 0
foreach ($cred in $matched) {
    $label = "$($cred.TargetName) [$(Get-CredTypeName $cred.Type)]"
    if ($PSCmdlet.ShouldProcess($label, 'Remove Windows credential')) {
        try {
            [RedSalamander.CredentialManager]::Delete($cred.TargetName, $cred.Type)
            Write-Host "  removed  $($cred.TargetName)" -ForegroundColor Green
            $removed++
        } catch {
            Write-Warning "  failed   $($cred.TargetName): $($_.Exception.Message)"
            $failed++
        }
    }
}

$isWhatIf = $WhatIfPreference -eq 'Continue'
if ($isWhatIf) {
    Write-Host "Would remove $($matched.Count) credential(s). Re-run without -WhatIf to apply." -ForegroundColor Yellow
} else {
    $color = if ($failed -gt 0) { 'Yellow' } else { 'Green' }
    Write-Host ("Removed {0} credential(s). Failed: {1}." -f $removed, $failed) -ForegroundColor $color
}
