<#
.SYNOPSIS
Runs one command while Windows Defender Firewall blocks outbound traffic from
every program on the evidence runner.

.DESCRIPTION
The provider creates one uniquely named explicit outbound Block rule whose
application, package, profile, interface, service, address, port, and protocol
filters are all unrestricted. This host-wide scope covers the root command and
all descendants, including executables generated after isolation starts.

The rule temporarily disconnects every process on the host. The caller must
therefore opt in with -ExclusiveDisposableRunner and run only on an exclusive,
disposable evidence runner. A bounded command timeout, exact rule ownership
checks, before/blocked/restored external canaries, authenticated-bypass scans,
and exact cleanup verification make any incomplete setup or cleanup a failed
attestation. Abrupt host termination can strand a PersistentStore rule, which
is another reason non-disposable or shared machines are forbidden.

An explicit block normally takes precedence over conflicting allow rules for
outbound traffic. Authenticated allow-bypass rules are the exception, so the
provider rejects any enabled outbound rule with OverrideBlockRules enabled.

This provider intentionally does not use New-NetFirewallRule -LocalUser.
Windows documents that filter as an authorized-user list for AppContainers,
not as a desktop-user scope for an arbitrary process tree.

.LINK
https://learn.microsoft.com/windows/security/operating-system-security/network-security/windows-firewall/rules

.LINK
https://learn.microsoft.com/powershell/module/netsecurity/new-netfirewallrule
#>
[CmdletBinding()]
param(
    [string]$FilePath = '',

    [string[]]$ArgumentList = @(),

    [string]$CanaryExecutablePath = '',

    [string]$WorkingDirectory = (Get-Location).Path,

    [string]$CanaryAddress = '1.1.1.1',

    [ValidateRange(1, 65535)]
    [int]$CanaryPort = 443,

    [ValidateRange(250, 30000)]
    [int]$CanaryTimeoutMilliseconds = 5000,

    [ValidateRange(1, 7200)]
    [int]$CommandTimeoutSeconds = 3600,

    [switch]$ExclusiveDisposableRunner,

    [switch]$LoadOnly
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$sanitizedEnvironmentScript = Join-Path (Split-Path -Parent $PSScriptRoot) 'SanitizedEnvironment.ps1'
if (-not (Test-Path -LiteralPath $sanitizedEnvironmentScript -PathType Leaf)) {
    throw "Contained-process helper not found: $sanitizedEnvironmentScript"
}
. $sanitizedEnvironmentScript

function Test-RSAdministrator {
    if ([System.Environment]::OSVersion.Platform -ne [System.PlatformID]::Win32NT) {
        return $false
    }

    $identity = [System.Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = [System.Security.Principal.WindowsPrincipal]::new($identity)
    return $principal.IsInRole([System.Security.Principal.WindowsBuiltInRole]::Administrator)
}

function Resolve-RSExactExecutablePath {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path,

        [Parameter(Mandatory = $true)]
        [string]$ParameterName
    )

    if ([string]::IsNullOrWhiteSpace($Path)) {
        throw "$ParameterName must name one executable by its exact canonical path."
    }
    if (-not [System.IO.Path]::IsPathFullyQualified($Path)) {
        throw "$ParameterName must be fully qualified: $Path"
    }

    $resolvedPaths = @(Resolve-Path -LiteralPath $Path -ErrorAction Stop)
    if ($resolvedPaths.Count -ne 1 -or
        $resolvedPaths[0].Provider.Name -ne 'FileSystem' -or
        -not (Test-Path -LiteralPath $resolvedPaths[0].ProviderPath -PathType Leaf)) {
        throw "$ParameterName must resolve to exactly one existing file: $Path"
    }

    $canonicalPath = [System.IO.Path]::GetFullPath($resolvedPaths[0].ProviderPath)
    if (-not [string]::Equals($Path, $canonicalPath, [System.StringComparison]::OrdinalIgnoreCase)) {
        throw "$ParameterName must already be the exact canonical path '$canonicalPath'; aliases, dot segments, and short names are not accepted."
    }
    if (-not [string]::Equals(
            [System.IO.Path]::GetExtension($canonicalPath),
            '.exe',
            [System.StringComparison]::OrdinalIgnoreCase)) {
        throw "$ParameterName must name an .exe file: $canonicalPath"
    }

    return $canonicalPath
}

function Resolve-RSExactDirectoryPath {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    if ([string]::IsNullOrWhiteSpace($Path) -or
        -not [System.IO.Path]::IsPathFullyQualified($Path)) {
        throw "WorkingDirectory must be an existing fully qualified directory: $Path"
    }

    $resolvedPaths = @(Resolve-Path -LiteralPath $Path -ErrorAction Stop)
    if ($resolvedPaths.Count -ne 1 -or
        $resolvedPaths[0].Provider.Name -ne 'FileSystem' -or
        -not (Test-Path -LiteralPath $resolvedPaths[0].ProviderPath -PathType Container)) {
        throw "WorkingDirectory must resolve to exactly one existing directory: $Path"
    }

    return [System.IO.Path]::GetFullPath($resolvedPaths[0].ProviderPath)
}

function Test-RSPublicIpv4Address {
    param(
        [Parameter(Mandatory = $true)]
        [System.Net.IPAddress]$Address
    )

    if ($Address.AddressFamily -ne [System.Net.Sockets.AddressFamily]::InterNetwork) {
        return $false
    }

    $octets = $Address.GetAddressBytes()
    if ($octets[0] -eq 0 -or
        $octets[0] -eq 10 -or
        $octets[0] -eq 127 -or
        $octets[0] -ge 224 -or
        ($octets[0] -eq 100 -and $octets[1] -ge 64 -and $octets[1] -le 127) -or
        ($octets[0] -eq 169 -and $octets[1] -eq 254) -or
        ($octets[0] -eq 172 -and $octets[1] -ge 16 -and $octets[1] -le 31) -or
        ($octets[0] -eq 192 -and $octets[1] -eq 168) -or
        ($octets[0] -eq 192 -and $octets[1] -eq 0 -and $octets[2] -in @(0, 2)) -or
        ($octets[0] -eq 198 -and $octets[1] -in @(18, 19)) -or
        ($octets[0] -eq 198 -and $octets[1] -eq 51 -and $octets[2] -eq 100) -or
        ($octets[0] -eq 203 -and $octets[1] -eq 0 -and $octets[2] -eq 113)) {
        return $false
    }

    return $true
}

function Resolve-RSCanaryAddress {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Address
    )

    $parsedAddress = $null
    if (-not [System.Net.IPAddress]::TryParse($Address, [ref]$parsedAddress) -or
        -not (Test-RSPublicIpv4Address -Address $parsedAddress)) {
        throw "CanaryAddress must be a literal public IPv4 address, not DNS, loopback, private, link-local, documentation, or multicast space: $Address"
    }

    return $parsedAddress
}

function Get-RSSha256 {
    param(
        [Parameter(Mandatory = $true)]
        [AllowEmptyCollection()]
        [byte[]]$Bytes
    )

    $hasher = [System.Security.Cryptography.SHA256]::Create()
    try {
        $digest = $hasher.ComputeHash($Bytes)
        return (($digest | ForEach-Object { $_.ToString('x2') }) -join '')
    }
    finally {
        $hasher.Dispose()
    }
}

function Get-RSFileIdentity {
    param(
        [Parameter(Mandatory = $true)]
        [string]$CanonicalPath
    )

    $file = [System.IO.FileInfo]::new($CanonicalPath)
    return [pscustomobject]@{
        path = $CanonicalPath
        sha256 = (Get-FileHash -LiteralPath $CanonicalPath -Algorithm SHA256).Hash.ToLowerInvariant()
        sizeBytes = $file.Length.ToString([System.Globalization.CultureInfo]::InvariantCulture)
    }
}

function Assert-RSFirewallProviderAvailable {
    $requiredCommands = @(
        'New-NetFirewallRule',
        'Get-NetFirewallRule',
        'Get-NetFirewallApplicationFilter',
        'Get-NetFirewallAddressFilter',
        'Get-NetFirewallPortFilter',
        'Get-NetFirewallServiceFilter',
        'Get-NetFirewallInterfaceFilter',
        'Get-NetFirewallInterfaceTypeFilter',
        'Get-NetFirewallSecurityFilter',
        'Remove-NetFirewallRule',
        'Get-NetFirewallProfile'
    )
    foreach ($commandName in $requiredCommands) {
        if ($null -eq (Get-Command $commandName -ErrorAction SilentlyContinue)) {
            throw "The runner-enforced network provider requires the NetSecurity command '$commandName'."
        }
    }

    $service = Get-Service -Name MpsSvc -ErrorAction Stop
    if ($service.Status -ne [System.ServiceProcess.ServiceControllerStatus]::Running) {
        throw 'Windows Defender Firewall service MpsSvc must be running.'
    }

    $profiles = @(Get-NetFirewallProfile -PolicyStore ActiveStore -ErrorAction Stop)
    $expectedNames = @('Domain', 'Private', 'Public')
    foreach ($profileName in $expectedNames) {
        $matching = @($profiles | Where-Object { $_.Name -eq $profileName })
        if ($matching.Count -ne 1 -or $matching[0].Enabled -ne $true) {
            throw "Windows Defender Firewall profile '$profileName' must be enabled in ActiveStore."
        }
    }

    return [pscustomobject]@{
        service = 'MpsSvc'
        serviceStatus = 'Running'
        policyStore = 'PersistentStore'
        effectivePolicyStore = 'ActiveStore'
        enabledProfiles = $expectedNames
    }
}

function Get-RSFirewallRulesByExactName {
    param(
        [Parameter(Mandatory = $true)]
        [ValidateSet('PersistentStore', 'ActiveStore')]
        [string]$PolicyStore,

        [Parameter(Mandatory = $true)]
        [string]$Name
    )

    try {
        return @(Get-NetFirewallRule `
                -PolicyStore $PolicyStore `
                -Name $Name `
                -ErrorAction Stop)
    }
    catch {
        if ($_.FullyQualifiedErrorId.StartsWith(
                'CmdletizationQuery_NotFound_InstanceID,',
                [System.StringComparison]::Ordinal)) {
            return
        }

        throw
    }
}

function New-RSOutboundBlockRule {
    param(
        [Parameter(Mandatory = $true)]
        [psobject]$Rule
    )

    $existingPersistent = @(Get-RSFirewallRulesByExactName `
            -PolicyStore PersistentStore `
            -Name $Rule.name)
    $existingActive = @(Get-RSFirewallRulesByExactName `
            -PolicyStore ActiveStore `
            -Name $Rule.name)
    if ($existingPersistent.Count -ne 0 -or $existingActive.Count -ne 0) {
        throw "Refusing to replace pre-existing firewall rule '$($Rule.name)'."
    }

    New-NetFirewallRule `
        -PolicyStore PersistentStore `
        -Name $Rule.name `
        -DisplayName $Rule.displayName `
        -Group $Rule.group `
        -Description $Rule.description `
        -Enabled True `
        -Direction Outbound `
        -Action Block `
        -Profile Any `
        -Protocol Any `
        -ErrorAction Stop | Out-Null
}

function Get-RSFirewallApplicationFilters {
    param(
        [Parameter(Mandatory = $true)]
        [object]$Rule,

        [Parameter(Mandatory = $true)]
        [ValidateSet('PersistentStore', 'ActiveStore')]
        [string]$PolicyStore
    )

    return @(Get-NetFirewallApplicationFilter `
            -PolicyStore $PolicyStore `
            -AssociatedNetFirewallRule $Rule `
            -ErrorAction Stop)
}

function Get-RSFirewallAddressFilters {
    param(
        [Parameter(Mandatory = $true)]
        [object]$Rule,

        [Parameter(Mandatory = $true)]
        [ValidateSet('PersistentStore', 'ActiveStore')]
        [string]$PolicyStore
    )

    return @(Get-NetFirewallAddressFilter `
            -PolicyStore $PolicyStore `
            -AssociatedNetFirewallRule $Rule `
            -ErrorAction Stop)
}

function Get-RSFirewallPortFilters {
    param(
        [Parameter(Mandatory = $true)]
        [object]$Rule,

        [Parameter(Mandatory = $true)]
        [ValidateSet('PersistentStore', 'ActiveStore')]
        [string]$PolicyStore
    )

    return @(Get-NetFirewallPortFilter `
            -PolicyStore $PolicyStore `
            -AssociatedNetFirewallRule $Rule `
            -ErrorAction Stop)
}

function Get-RSFirewallServiceFilters {
    param(
        [Parameter(Mandatory = $true)]
        [object]$Rule,

        [Parameter(Mandatory = $true)]
        [ValidateSet('PersistentStore', 'ActiveStore')]
        [string]$PolicyStore
    )

    return @(Get-NetFirewallServiceFilter `
            -PolicyStore $PolicyStore `
            -AssociatedNetFirewallRule $Rule `
            -ErrorAction Stop)
}

function Get-RSFirewallInterfaceFilters {
    param(
        [Parameter(Mandatory = $true)]
        [object]$Rule,

        [Parameter(Mandatory = $true)]
        [ValidateSet('PersistentStore', 'ActiveStore')]
        [string]$PolicyStore
    )

    return @(Get-NetFirewallInterfaceFilter `
            -PolicyStore $PolicyStore `
            -AssociatedNetFirewallRule $Rule `
            -ErrorAction Stop)
}

function Get-RSFirewallInterfaceTypeFilters {
    param(
        [Parameter(Mandatory = $true)]
        [object]$Rule,

        [Parameter(Mandatory = $true)]
        [ValidateSet('PersistentStore', 'ActiveStore')]
        [string]$PolicyStore
    )

    return @(Get-NetFirewallInterfaceTypeFilter `
            -PolicyStore $PolicyStore `
            -AssociatedNetFirewallRule $Rule `
            -ErrorAction Stop)
}

function Get-RSFirewallSecurityFilters {
    param(
        [Parameter(Mandatory = $true)]
        [object]$Rule,

        [Parameter(Mandatory = $true)]
        [ValidateSet('PersistentStore', 'ActiveStore')]
        [string]$PolicyStore
    )

    return @(Get-NetFirewallSecurityFilter `
            -PolicyStore $PolicyStore `
            -AssociatedNetFirewallRule $Rule `
            -ErrorAction Stop)
}

function Assert-RSNoOutboundBlockBypass {
    $allowRules = @(Get-NetFirewallRule `
            -PolicyStore ActiveStore `
            -Enabled True `
            -Direction Outbound `
            -Action Allow `
            -ErrorAction Stop)
    $bypassRuleNames = [System.Collections.Generic.List[string]]::new()
    foreach ($allowRule in $allowRules) {
        $securityFilters = @(Get-RSFirewallSecurityFilters `
                -PolicyStore ActiveStore `
                -Rule $allowRule)
        if ($securityFilters.Count -ne 1) {
            throw "Enabled outbound allow rule '$($allowRule.Name)' has $($securityFilters.Count) security filters; block-bypass state is ambiguous."
        }
        if ([string]::Equals(
                [string]$securityFilters[0].OverrideBlockRules,
                'True',
                [System.StringComparison]::OrdinalIgnoreCase)) {
            [void]$bypassRuleNames.Add([string]$allowRule.Name)
        }
    }

    if ($bypassRuleNames.Count -ne 0) {
        throw "Enabled outbound authenticated-bypass rules can override a block rule: $($bypassRuleNames -join ', ')."
    }

    return [pscustomobject]@{
        enabledOutboundAllowRulesScanned = $allowRules.Count
        outboundBlockBypassRules = 0
    }
}

function Assert-RSOutboundBlockRule {
    param(
        [Parameter(Mandatory = $true)]
        [psobject]$Rule
    )

    $lastFailure = ''
    for ($attempt = 0; $attempt -lt 30; $attempt++) {
        $rules = @(Get-RSFirewallRulesByExactName `
                -PolicyStore ActiveStore `
                -Name $Rule.name)
        if ($rules.Count -eq 1) {
            $actualRule = $rules[0]
            $filters = @(Get-RSFirewallApplicationFilters `
                    -PolicyStore ActiveStore `
                    -Rule $actualRule)
            $addressFilters = @(Get-RSFirewallAddressFilters `
                    -PolicyStore ActiveStore `
                    -Rule $actualRule)
            $portFilters = @(Get-RSFirewallPortFilters `
                    -PolicyStore ActiveStore `
                    -Rule $actualRule)
            $serviceFilters = @(Get-RSFirewallServiceFilters `
                    -PolicyStore ActiveStore `
                    -Rule $actualRule)
            $interfaceFilters = @(Get-RSFirewallInterfaceFilters `
                    -PolicyStore ActiveStore `
                    -Rule $actualRule)
            $interfaceTypeFilters = @(Get-RSFirewallInterfaceTypeFilters `
                    -PolicyStore ActiveStore `
                    -Rule $actualRule)
            $securityFilters = @(Get-RSFirewallSecurityFilters `
                    -PolicyStore ActiveStore `
                    -Rule $actualRule)
            if ($filters.Count -eq 1 -and
                $addressFilters.Count -eq 1 -and
                $portFilters.Count -eq 1 -and
                $serviceFilters.Count -eq 1 -and
                $interfaceFilters.Count -eq 1 -and
                $interfaceTypeFilters.Count -eq 1 -and
                $securityFilters.Count -eq 1 -and
                [string]::Equals($actualRule.Name, $Rule.name, [System.StringComparison]::Ordinal) -and
                [string]::Equals($actualRule.Group, $Rule.group, [System.StringComparison]::Ordinal) -and
                [string]::Equals([string]$actualRule.Enabled, 'True', [System.StringComparison]::OrdinalIgnoreCase) -and
                [string]::Equals([string]$actualRule.Direction, 'Outbound', [System.StringComparison]::OrdinalIgnoreCase) -and
                [string]::Equals([string]$actualRule.Action, 'Block', [System.StringComparison]::OrdinalIgnoreCase) -and
                [string]::Equals([string]$actualRule.Profile, 'Any', [System.StringComparison]::OrdinalIgnoreCase) -and
                [string]::Equals([string]$filters[0].Program, 'Any', [System.StringComparison]::OrdinalIgnoreCase) -and
                ([string]::IsNullOrEmpty([string]$filters[0].Package) -or
                    [string]::Equals([string]$filters[0].Package, 'Any', [System.StringComparison]::OrdinalIgnoreCase)) -and
                [string]::Equals([string]$addressFilters[0].LocalAddress, 'Any', [System.StringComparison]::OrdinalIgnoreCase) -and
                [string]::Equals([string]$addressFilters[0].RemoteAddress, 'Any', [System.StringComparison]::OrdinalIgnoreCase) -and
                [string]::Equals([string]$portFilters[0].Protocol, 'Any', [System.StringComparison]::OrdinalIgnoreCase) -and
                [string]::Equals([string]$portFilters[0].LocalPort, 'Any', [System.StringComparison]::OrdinalIgnoreCase) -and
                [string]::Equals([string]$portFilters[0].RemotePort, 'Any', [System.StringComparison]::OrdinalIgnoreCase) -and
                [string]::Equals([string]$serviceFilters[0].Service, 'Any', [System.StringComparison]::OrdinalIgnoreCase) -and
                [string]::Equals([string]$interfaceFilters[0].InterfaceAlias, 'Any', [System.StringComparison]::OrdinalIgnoreCase) -and
                [string]::Equals([string]$interfaceTypeFilters[0].InterfaceType, 'Any', [System.StringComparison]::OrdinalIgnoreCase) -and
                [string]::Equals([string]$securityFilters[0].OverrideBlockRules, 'False', [System.StringComparison]::OrdinalIgnoreCase)) {
                return
            }

            $lastFailure = 'rule/filter properties did not match the requested group, all programs/packages/profiles/interfaces/services/addresses/ports/protocols, no-bypass security state, and outbound block action'
        }
        else {
            $lastFailure = "ActiveStore contained $($rules.Count) matching rules"
        }

        Start-Sleep -Milliseconds 100
    }

    throw "Firewall rule '$($Rule.name)' was not verified in ActiveStore: $lastFailure."
}

function Remove-RSExactFirewallRules {
    param(
        [Parameter(Mandatory = $true)]
        [psobject[]]$Rules
    )

    $cleanupFailures = [System.Collections.Generic.List[string]]::new()
    foreach ($rule in $Rules) {
        try {
            $persistentRules = @(Get-RSFirewallRulesByExactName `
                    -PolicyStore PersistentStore `
                    -Name $rule.name)
            if ($persistentRules.Count -gt 1) {
                throw "PersistentStore contains $($persistentRules.Count) rules with the exact generated name."
            }

            if ($persistentRules.Count -eq 1) {
                $persistentRule = $persistentRules[0]
                if (-not [string]::Equals($persistentRule.Name, $rule.name, [System.StringComparison]::Ordinal) -or
                    -not [string]::Equals($persistentRule.DisplayName, $rule.displayName, [System.StringComparison]::Ordinal) -or
                    -not [string]::Equals($persistentRule.Group, $rule.group, [System.StringComparison]::Ordinal) -or
                    -not [string]::Equals($persistentRule.Description, $rule.description, [System.StringComparison]::Ordinal) -or
                    -not [string]::Equals([string]$persistentRule.Direction, 'Outbound', [System.StringComparison]::OrdinalIgnoreCase) -or
                    -not [string]::Equals([string]$persistentRule.Action, 'Block', [System.StringComparison]::OrdinalIgnoreCase)) {
                    throw 'The generated rule identity changed; refusing a broad or ambiguous removal.'
                }

                $filters = @(Get-RSFirewallApplicationFilters `
                        -PolicyStore PersistentStore `
                        -Rule $persistentRule)
                if ($filters.Count -ne 1 -or
                    -not [string]::Equals([string]$filters[0].Program, 'Any', [System.StringComparison]::OrdinalIgnoreCase) -or
                    (-not [string]::IsNullOrEmpty([string]$filters[0].Package) -and
                        -not [string]::Equals([string]$filters[0].Package, 'Any', [System.StringComparison]::OrdinalIgnoreCase))) {
                    throw 'The generated rule identity changed; refusing a broad or ambiguous removal.'
                }

                Remove-NetFirewallRule `
                    -PolicyStore PersistentStore `
                    -Name $rule.name `
                    -Confirm:$false `
                    -ErrorAction Stop
            }
        }
        catch {
            [void]$cleanupFailures.Add("$($rule.name): $($_.Exception.Message)")
        }
    }

    foreach ($rule in $Rules) {
        $removed = $false
        for ($attempt = 0; $attempt -lt 30; $attempt++) {
            $persistentRules = @(Get-RSFirewallRulesByExactName `
                    -PolicyStore PersistentStore `
                    -Name $rule.name)
            $activeRules = @(Get-RSFirewallRulesByExactName `
                    -PolicyStore ActiveStore `
                    -Name $rule.name)
            if ($persistentRules.Count -eq 0 -and $activeRules.Count -eq 0) {
                $removed = $true
                break
            }

            Start-Sleep -Milliseconds 100
        }

        if (-not $removed) {
            [void]$cleanupFailures.Add("$($rule.name): exact rule remained after cleanup verification")
        }
    }

    if ($cleanupFailures.Count -gt 0) {
        throw "Exact firewall-rule cleanup failed: $($cleanupFailures -join '; ')"
    }
}

function Invoke-RSCapturedProcess {
    param(
        [Parameter(Mandatory = $true)]
        [string]$CanonicalFilePath,

        [string[]]$Arguments = @(),

        [Parameter(Mandatory = $true)]
        [string]$CanonicalWorkingDirectory,

        [Parameter(Mandatory = $true)]
        [int]$TimeoutMilliseconds
    )

    $startInfo = New-RSProcessStartInfo `
        -FilePath $CanonicalFilePath `
        -Arguments $Arguments `
        -WorkingDirectory $CanonicalWorkingDirectory
    $startInfo.RedirectStandardOutput = $true
    $startInfo.RedirectStandardError = $true
    $startInfo.CreateNoWindow = $true

    $process = $null
    $startedUtc = [System.DateTime]::UtcNow
    $stopwatch = [System.Diagnostics.Stopwatch]::StartNew()
    try {
        $process = Start-RSContainedProcess -ProcessStartInfo $startInfo
        if (-not $process.IsInContainmentJob) {
            throw "Exact executable '$CanonicalFilePath' did not enter the kill-on-close containment job."
        }

        $stdoutTask = $process.StandardOutput.ReadToEndAsync()
        $stderrTask = $process.StandardError.ReadToEndAsync()
        $exited = $process.WaitForExit($TimeoutMilliseconds)
        $exitCode = if ($exited) { $process.ExitCode } else { $null }

        $streamTasks = [System.Threading.Tasks.Task[]]@($stdoutTask, $stderrTask)
        $streamsCompleted = $false
        if ($exited) {
            $streamsCompleted = [System.Threading.Tasks.Task]::WaitAll($streamTasks, 5000)
        }

        $processTreeForcedClosed = (-not $exited -or -not $streamsCompleted)
        if ($processTreeForcedClosed) {
            Close-RSContainedProcess -Process $process
            $process = $null
            [void][System.Threading.Tasks.Task]::WaitAll($streamTasks, 5000)
        }

        $outputComplete = ($stdoutTask.IsCompletedSuccessfully -and $stderrTask.IsCompletedSuccessfully)
        $stdout = if ($stdoutTask.IsCompletedSuccessfully) {
            $stdoutTask.GetAwaiter().GetResult()
        }
        else {
            ''
        }
        $stderr = if ($stderrTask.IsCompletedSuccessfully) {
            $stderrTask.GetAwaiter().GetResult()
        }
        else {
            ''
        }
        $stopwatch.Stop()
        $completedUtc = [System.DateTime]::UtcNow
        $encoding = [System.Text.UTF8Encoding]::new($false)

        return [pscustomobject]@{
            exitCode = $exitCode
            timedOut = -not $exited
            processTreeContained = $true
            processTreeForcedClosed = $processTreeForcedClosed
            outputComplete = $outputComplete
            stdout = $stdout
            stderr = $stderr
            stdoutSha256 = Get-RSSha256 -Bytes $encoding.GetBytes($stdout)
            stderrSha256 = Get-RSSha256 -Bytes $encoding.GetBytes($stderr)
            startedUtc = $startedUtc.ToString('o', [System.Globalization.CultureInfo]::InvariantCulture)
            completedUtc = $completedUtc.ToString('o', [System.Globalization.CultureInfo]::InvariantCulture)
            durationMilliseconds = $stopwatch.ElapsedMilliseconds.ToString([System.Globalization.CultureInfo]::InvariantCulture)
        }
    }
    finally {
        $stopwatch.Stop()
        Close-RSContainedProcess -Process $process
    }
}

function Invoke-RSExternalConnectCanary {
    param(
        [Parameter(Mandatory = $true)]
        [string]$CanonicalCanaryExecutablePath,

        [Parameter(Mandatory = $true)]
        [System.Net.IPAddress]$Address,

        [Parameter(Mandatory = $true)]
        [int]$Port,

        [Parameter(Mandatory = $true)]
        [int]$TimeoutMilliseconds,

        [Parameter(Mandatory = $true)]
        [string]$CanonicalWorkingDirectory
    )

    $leaf = [System.IO.Path]::GetFileName($CanonicalCanaryExecutablePath)
    if (-not [string]::Equals($leaf, 'pwsh.exe', [System.StringComparison]::OrdinalIgnoreCase) -and
        -not [string]::Equals($leaf, 'powershell.exe', [System.StringComparison]::OrdinalIgnoreCase)) {
        throw "CanaryExecutablePath must name pwsh.exe or powershell.exe so the provider can run its fixed TCP canary: $CanonicalCanaryExecutablePath"
    }

    $addressLiteral = $Address.ToString()
    $canaryScript = @"
`$ErrorActionPreference = 'Stop'
`$client = [System.Net.Sockets.TcpClient]::new([System.Net.Sockets.AddressFamily]::InterNetwork)
try {
    `$pending = `$client.BeginConnect('$addressLiteral', $Port, `$null, `$null)
    if (-not `$pending.AsyncWaitHandle.WaitOne($TimeoutMilliseconds)) {
        [Console]::Error.Write('connect-timeout')
        exit 28
    }
    `$client.EndConnect(`$pending)
    if (-not `$client.Connected) {
        [Console]::Error.Write('connect-not-connected')
        exit 29
    }
    [Console]::Out.Write('connected')
    exit 0
}
catch {
    [Console]::Error.Write(`$_.Exception.GetType().FullName)
    exit 30
}
finally {
    `$client.Dispose()
}
"@

    $result = Invoke-RSCapturedProcess `
        -CanonicalFilePath $CanonicalCanaryExecutablePath `
        -Arguments @(
            '-NoLogo',
            '-NoProfile',
            '-NonInteractive',
            '-ExecutionPolicy',
            'Bypass',
            '-Command',
            $canaryScript
        ) `
        -CanonicalWorkingDirectory $CanonicalWorkingDirectory `
        -TimeoutMilliseconds ($TimeoutMilliseconds + 5000)

    return [pscustomobject]@{
        connected = ($result.exitCode -eq 0 -and
            -not $result.timedOut -and
            -not $result.processTreeForcedClosed -and
            $result.outputComplete)
        exitCode = $result.exitCode
        timedOut = $result.timedOut
        stdout = $result.stdout
        stderr = $result.stderr
        durationMilliseconds = $result.durationMilliseconds
    }
}

function Invoke-RSRootCommand {
    param(
        [Parameter(Mandatory = $true)]
        [string]$CanonicalFilePath,

        [string[]]$Arguments = @(),

        [Parameter(Mandatory = $true)]
        [string]$CanonicalWorkingDirectory,

        [Parameter(Mandatory = $true)]
        [int]$TimeoutSeconds
    )

    return Invoke-RSCapturedProcess `
        -CanonicalFilePath $CanonicalFilePath `
        -Arguments $Arguments `
        -CanonicalWorkingDirectory $CanonicalWorkingDirectory `
        -TimeoutMilliseconds ($TimeoutSeconds * 1000)
}

function New-RSFirewallRulePlan {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Group
    )

    return @([pscustomobject]@{
            name = "$Group.AllPrograms"
            displayName = 'RedSalamander Terminal Gate 0 whole-runner isolation'
            group = $Group
            description = "Ephemeral all-program outbound block owned by $Group"
        })
}

function Invoke-RSIsolatedProcess {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$FilePath,

        [string[]]$ArgumentList = @(),

        [Parameter(Mandatory = $true)]
        [string]$CanaryExecutablePath,

        [string]$WorkingDirectory = (Get-Location).Path,

        [string]$CanaryAddress = '1.1.1.1',

        [ValidateRange(1, 65535)]
        [int]$CanaryPort = 443,

        [ValidateRange(250, 30000)]
        [int]$CanaryTimeoutMilliseconds = 5000,

        [ValidateRange(1, 7200)]
        [int]$CommandTimeoutSeconds = 3600,

        [switch]$ExclusiveDisposableRunner
    )

    if (-not $ExclusiveDisposableRunner) {
        throw 'Host-wide outbound isolation requires -ExclusiveDisposableRunner; this provider may run only on an exclusive disposable evidence runner.'
    }
    if (-not (Test-RSAdministrator)) {
        throw 'Runner-enforced outbound isolation requires an elevated Windows administrator process.'
    }

    $canonicalFilePath = Resolve-RSExactExecutablePath -Path $FilePath -ParameterName 'FilePath'
    $canonicalCanaryPath = Resolve-RSExactExecutablePath `
        -Path $CanaryExecutablePath `
        -ParameterName 'CanaryExecutablePath'
    $canonicalWorkingDirectory = Resolve-RSExactDirectoryPath -Path $WorkingDirectory
    $publicCanaryAddress = Resolve-RSCanaryAddress -Address $CanaryAddress

    $canonicalPrograms = [System.Collections.Generic.List[string]]::new()
    [void]$canonicalPrograms.Add($canonicalFilePath)
    if (-not [string]::Equals(
            $canonicalCanaryPath,
            $canonicalFilePath,
            [System.StringComparison]::OrdinalIgnoreCase)) {
        [void]$canonicalPrograms.Add($canonicalCanaryPath)
    }

    $firewallState = Assert-RSFirewallProviderAvailable
    $bypassScanBefore = Assert-RSNoOutboundBlockBypass
    $runGuid = [System.Guid]::NewGuid().ToString('N')
    $ruleGroup = "RedSalamander.TerminalEngine.Gate0.$runGuid"
    $rulePlan = @(New-RSFirewallRulePlan -Group $ruleGroup)
    $programIdentities = @($canonicalPrograms | ForEach-Object {
            Get-RSFileIdentity -CanonicalPath $_
        })
    $rootIdentity = @($programIdentities | Where-Object {
            [string]::Equals($_.path, $canonicalFilePath, [System.StringComparison]::OrdinalIgnoreCase)
        })[0]
    $canaryIdentity = Get-RSFileIdentity -CanonicalPath $canonicalCanaryPath
    $failures = [System.Collections.Generic.List[object]]::new()
    $preCanary = $null
    $blockedCanary = $null
    $restoredCanary = $null
    $rootResult = $null
    $ruleSetupStarted = $false
    $ruleSetupVerified = $false
    $ruleRemovalVerified = $false
    $bypassScanDuring = $null
    $startedUtc = [System.DateTime]::UtcNow

    try {
        $preCanary = Invoke-RSExternalConnectCanary `
            -CanonicalCanaryExecutablePath $canonicalCanaryPath `
            -Address $publicCanaryAddress `
            -Port $CanaryPort `
            -TimeoutMilliseconds $CanaryTimeoutMilliseconds `
            -CanonicalWorkingDirectory $canonicalWorkingDirectory
        if (-not $preCanary.connected) {
            throw "The before-isolation external canary could not connect (exit=$($preCanary.exitCode)); enforcement cannot be distinguished from an already-offline runner."
        }

        $ruleSetupStarted = $true
        foreach ($rule in $rulePlan) {
            New-RSOutboundBlockRule -Rule $rule
        }
        foreach ($rule in $rulePlan) {
            Assert-RSOutboundBlockRule -Rule $rule
        }
        $bypassScanDuring = Assert-RSNoOutboundBlockBypass
        $ruleSetupVerified = $true

        $blockedCanary = Invoke-RSExternalConnectCanary `
            -CanonicalCanaryExecutablePath $canonicalCanaryPath `
            -Address $publicCanaryAddress `
            -Port $CanaryPort `
            -TimeoutMilliseconds $CanaryTimeoutMilliseconds `
            -CanonicalWorkingDirectory $canonicalWorkingDirectory
        if ($blockedCanary.connected) {
            throw 'The during-isolation external canary connected despite the verified host-wide outbound block rule.'
        }

        $rootResult = Invoke-RSRootCommand `
            -CanonicalFilePath $canonicalFilePath `
            -Arguments $ArgumentList `
            -CanonicalWorkingDirectory $canonicalWorkingDirectory `
            -TimeoutSeconds $CommandTimeoutSeconds
    }
    catch {
        $stage = if (-not $preCanary) {
            'setup-canary-before'
        }
        elseif (-not $preCanary.connected) {
            'setup-canary-before'
        }
        elseif (-not $ruleSetupVerified) {
            'firewall-setup'
        }
        elseif (-not $blockedCanary -or $blockedCanary.connected) {
            'setup-canary-during'
        }
        else {
            'root-command'
        }
        [void]$failures.Add([pscustomobject]@{
                stage = $stage
                message = $_.Exception.Message
            })
    }
    finally {
        if ($ruleSetupStarted) {
            try {
                Remove-RSExactFirewallRules -Rules $rulePlan
                $ruleRemovalVerified = $true
                $restoredCanary = Invoke-RSExternalConnectCanary `
                    -CanonicalCanaryExecutablePath $canonicalCanaryPath `
                    -Address $publicCanaryAddress `
                    -Port $CanaryPort `
                    -TimeoutMilliseconds $CanaryTimeoutMilliseconds `
                    -CanonicalWorkingDirectory $canonicalWorkingDirectory
                if (-not $restoredCanary.connected) {
                    throw "The after-cleanup external canary did not reconnect (exit=$($restoredCanary.exitCode))."
                }
            }
            catch {
                $ruleRemovalVerified = $false
                [void]$failures.Add([pscustomobject]@{
                        stage = 'cleanup'
                        message = $_.Exception.Message
                    })
            }
        }
        else {
            $ruleRemovalVerified = $true
        }

    }

    $programIdentitiesAfter = @()
    $programIdentityStable = $false
    try {
        $programIdentitiesAfter = @($canonicalPrograms | ForEach-Object {
                Get-RSFileIdentity -CanonicalPath $_
            })
        $programIdentityStable = ($programIdentitiesAfter.Count -eq $programIdentities.Count)
        if ($programIdentityStable) {
            for ($index = 0; $index -lt $programIdentities.Count; $index++) {
                $beforeIdentity = $programIdentities[$index]
                $afterIdentity = $programIdentitiesAfter[$index]
                if (-not [string]::Equals(
                        $beforeIdentity.path,
                        $afterIdentity.path,
                        [System.StringComparison]::OrdinalIgnoreCase) -or
                    -not [string]::Equals(
                        $beforeIdentity.sha256,
                        $afterIdentity.sha256,
                        [System.StringComparison]::Ordinal) -or
                    -not [string]::Equals(
                        $beforeIdentity.sizeBytes,
                        $afterIdentity.sizeBytes,
                        [System.StringComparison]::Ordinal)) {
                    $programIdentityStable = $false
                    break
                }
            }
        }

        if (-not $programIdentityStable) {
            [void]$failures.Add([pscustomobject]@{
                    stage = 'executable-identity'
                    message = 'The root or canary executable changed identity during the isolated run.'
                })
        }
    }
    catch {
        [void]$failures.Add([pscustomobject]@{
                stage = 'executable-identity'
                message = "Failed to revalidate root/canary executable identities: $($_.Exception.Message)"
            })
    }

    $completedUtc = [System.DateTime]::UtcNow
    $providerPassed = ($failures.Count -eq 0 -and
        $ruleSetupVerified -and
        $ruleRemovalVerified -and
        $null -ne $bypassScanBefore -and
        $bypassScanBefore.outboundBlockBypassRules -eq 0 -and
        $null -ne $bypassScanDuring -and
        $bypassScanDuring.outboundBlockBypassRules -eq 0 -and
        $programIdentityStable -and
        $null -ne $preCanary -and
        $preCanary.connected -and
        $null -ne $blockedCanary -and
        -not $blockedCanary.connected -and
        $null -ne $restoredCanary -and
        $restoredCanary.connected -and
        $null -ne $rootResult)

    return [pscustomobject]@{
        formatId = 'red-salamander-runner-network-isolation-attestation'
        schemaVersion = 1
        provider = 'windows-defender-firewall-host-wide-outbound-block'
        enforcement = 'runner-enforced'
        providerStatus = if ($providerPassed) { 'pass' } else { 'fail' }
        isolationVerified = $providerPassed
        rootCommandExecuted = ($null -ne $rootResult)
        runId = $runGuid
        startedUtc = $startedUtc.ToString('o', [System.Globalization.CultureInfo]::InvariantCulture)
        completedUtc = $completedUtc.ToString('o', [System.Globalization.CultureInfo]::InvariantCulture)
        firewall = $firewallState
        policyBypassScan = [pscustomobject]@{
            beforeIsolation = $bypassScanBefore
            duringIsolation = $bypassScanDuring
        }
        ruleGroup = $ruleGroup
        ruleNames = @($rulePlan | ForEach-Object { $_.name })
        ruleSetupVerified = $ruleSetupVerified
        ruleRemovalVerified = $ruleRemovalVerified
        programs = $programIdentities
        programsAfter = $programIdentitiesAfter
        programIdentityStable = $programIdentityStable
        canary = [pscustomobject]@{
            executable = $canaryIdentity
            address = $publicCanaryAddress.ToString()
            port = $CanaryPort
            beforeIsolation = $preCanary
            duringIsolation = $blockedCanary
            afterCleanup = $restoredCanary
        }
        rootCommand = [pscustomobject]@{
            executable = $rootIdentity
            argumentCount = $ArgumentList.Count
            result = $rootResult
        }
        scopeContract = [pscustomobject]@{
            hostWide = $true
            descendantCoverage = 'all-programs'
            exactProgramPathsOnly = $false
            completeExecutableClosureRequired = $false
            unlistedExecutablesAreBlocked = $true
            exclusiveDisposableRunnerAcknowledged = $true
            maximumCommandDurationSeconds = 7200
            proxyOrEnvironmentEvidenceAccepted = $false
        }
        failures = $failures.ToArray()
    }
}

if (-not $LoadOnly) {
    if ([string]::IsNullOrWhiteSpace($FilePath) -or
        [string]::IsNullOrWhiteSpace($CanaryExecutablePath)) {
        throw 'Direct invocation requires -FilePath and -CanaryExecutablePath.'
    }

    $attestation = Invoke-RSIsolatedProcess `
        -FilePath $FilePath `
        -ArgumentList $ArgumentList `
        -CanaryExecutablePath $CanaryExecutablePath `
        -WorkingDirectory $WorkingDirectory `
        -CanaryAddress $CanaryAddress `
        -CanaryPort $CanaryPort `
        -CanaryTimeoutMilliseconds $CanaryTimeoutMilliseconds `
        -CommandTimeoutSeconds $CommandTimeoutSeconds `
        -ExclusiveDisposableRunner:$ExclusiveDisposableRunner
    $attestation

    if ($attestation.providerStatus -ne 'pass') {
        $global:LASTEXITCODE = 74
    }
    elseif ($null -ne $attestation.rootCommand.result.exitCode) {
        $global:LASTEXITCODE = $attestation.rootCommand.result.exitCode
    }
}
