Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
$helperScript = Join-Path $repoRoot 'Tools\TerminalEngine\Invoke-RSIsolatedProcess.ps1'

Describe 'Terminal-engine runner network-isolation provider' {
    BeforeAll {
        . $helperScript -LoadOnly
        $script:pwshPath = [System.IO.Path]::GetFullPath(
            [System.Diagnostics.Process]::GetCurrentProcess().MainModule.FileName)
        $script:workingDirectory = [System.IO.Path]::GetFullPath($repoRoot)
    }

    Context 'input and source contracts' {
        It 'requires an elevated administrator process before touching the firewall provider' {
            Mock Test-RSAdministrator { return $false }
            Mock Assert-RSFirewallProviderAvailable { throw 'firewall provider must not be reached' }

            {
                Invoke-RSIsolatedProcess `
                    -FilePath $script:pwshPath `
                    -ArgumentList @('-NoProfile', '-Command', 'exit 0') `
                    -CanaryExecutablePath $script:pwshPath `
                    -WorkingDirectory $script:workingDirectory `
                    -ExclusiveDisposableRunner
            } | Should Throw 'Runner-enforced outbound isolation requires an elevated Windows administrator process.'

            Assert-MockCalled Assert-RSFirewallProviderAvailable -Times 0 -Scope It
        }

        It 'rejects host-wide mutation without explicit disposable-runner acknowledgement' {
            Mock Test-RSAdministrator { return $true }
            Mock Assert-RSFirewallProviderAvailable { throw 'firewall provider must not be reached' }

            {
                Invoke-RSIsolatedProcess `
                    -FilePath $script:pwshPath `
                    -ArgumentList @('-NoProfile', '-Command', 'exit 0') `
                    -CanaryExecutablePath $script:pwshPath `
                    -WorkingDirectory $script:workingDirectory
            } | Should Throw 'Host-wide outbound isolation requires -ExclusiveDisposableRunner; this provider may run only on an exclusive disposable evidence runner.'

            Assert-MockCalled Assert-RSFirewallProviderAvailable -Times 0 -Scope It
        }

        It 'rejects a fully qualified executable spelling that still contains a dot segment' {
            $pwshDirectory = Split-Path -Parent $script:pwshPath
            $pwshParent = Split-Path -Parent $pwshDirectory
            $noncanonicalPath = Join-Path $pwshParent (
                (Split-Path -Leaf $pwshDirectory) + '\..\'+
                (Split-Path -Leaf $pwshDirectory) + '\' +
                (Split-Path -Leaf $script:pwshPath))

            {
                Resolve-RSExactExecutablePath `
                    -Path $noncanonicalPath `
                    -ParameterName 'SyntheticPath'
            } | Should Throw "SyntheticPath must already be the exact canonical path '$($script:pwshPath)'; aliases, dot segments, and short names are not accepted."
        }

        It 'rejects private, loopback, documentation, and IPv6 canary addresses' {
            (Test-RSPublicIpv4Address -Address ([System.Net.IPAddress]::Parse('1.1.1.1'))) | Should Be $true
            (Test-RSPublicIpv4Address -Address ([System.Net.IPAddress]::Parse('10.0.0.1'))) | Should Be $false
            (Test-RSPublicIpv4Address -Address ([System.Net.IPAddress]::Parse('127.0.0.1'))) | Should Be $false
            (Test-RSPublicIpv4Address -Address ([System.Net.IPAddress]::Parse('192.0.2.1'))) | Should Be $false
            (Test-RSPublicIpv4Address -Address ([System.Net.IPAddress]::Parse('203.0.113.1'))) | Should Be $false
            (Test-RSPublicIpv4Address -Address ([System.Net.IPAddress]::Parse('2001:4860:4860::8888'))) | Should Be $false
        }

        It 'uses one all-program rule and never substitutes profile, proxy, or environment evidence' {
            $source = Get-Content -LiteralPath $helperScript -Raw

            $source | Should Match 'New-NetFirewallRule'
            $source | Should Match 'Remove-NetFirewallRule'
            $source | Should Match 'Get-NetFirewallApplicationFilter'
            $source | Should Match 'Get-NetFirewallAddressFilter'
            $source | Should Match 'Get-NetFirewallPortFilter'
            $source | Should Match 'Get-NetFirewallServiceFilter'
            $source | Should Match 'Get-NetFirewallInterfaceFilter'
            $source | Should Match 'Get-NetFirewallInterfaceTypeFilter'
            $source | Should Match 'PersistentStore'
            $source | Should Match 'ActiveStore'
            $source | Should Match "Program, 'Any'"
            $source | Should Match 'Get-NetFirewallSecurityFilter'
            $source | Should Match 'OverrideBlockRules'
            $source | Should Match 'ExclusiveDisposableRunner'
            $source | Should Match 'runner-enforced'
            $source | Should Not Match 'Set-NetFirewallProfile'
            $source | Should Not Match 'Remove-NetFirewallRule\s+.*-All'
            $source | Should Not Match 'HTTP_PROXY'
            $source | Should Not Match 'HTTPS_PROXY'
            $source | Should Not Match 'NO_PROXY'
        }

        It 'contains the process tree and kills a child that outlives the one root command' {
            $childPid = 0
            try {
                $command = @'
$child = Start-Process `
    -FilePath "$env:SystemRoot\System32\cmd.exe" `
    -ArgumentList @('/d', '/c', 'ping -n 60 127.0.0.1 >nul') `
    -PassThru `
    -WindowStyle Hidden
[Console]::Out.Write($child.Id)
exit 0
'@
                $result = Invoke-RSRootCommand `
                    -CanonicalFilePath $script:pwshPath `
                    -Arguments @('-NoProfile', '-NonInteractive', '-Command', $command) `
                    -CanonicalWorkingDirectory $script:workingDirectory `
                    -TimeoutSeconds 15

                $result.exitCode | Should Be 0
                $result.processTreeContained | Should Be $true
                $result.outputComplete | Should Be $true
                $childPid = [int]$result.stdout

                $childExited = $false
                for ($attempt = 0; $attempt -lt 40; $attempt++) {
                    if ($null -eq (Get-Process -Id $childPid -ErrorAction SilentlyContinue)) {
                        $childExited = $true
                        break
                    }
                    Start-Sleep -Milliseconds 25
                }
                $childExited | Should Be $true
            }
            finally {
                if ($childPid -gt 0) {
                    Stop-Process -Id $childPid -Force -ErrorAction SilentlyContinue
                }
            }
        }

        It 'hashes captured empty output deterministically' {
            (Get-RSSha256 -Bytes ([byte[]]@())) | Should Be 'e3b0c44298fc1c149afbf4c8996fb92427ae41e4649b934ca495991b7852b855'
        }

        It 'distinguishes an exact missing firewall-rule name from a provider query failure' {
            $missingName = 'RedSalamander.TerminalEngine.Gate0.' +
                [System.Guid]::NewGuid().ToString('N') +
                '.AllPrograms'

            @(Get-RSFirewallRulesByExactName `
                    -PolicyStore ActiveStore `
                    -Name $missingName).Count | Should Be 0

            Mock Get-NetFirewallRule { throw 'synthetic firewall query failure' }
            {
                Get-RSFirewallRulesByExactName `
                    -PolicyStore ActiveStore `
                    -Name $missingName
            } | Should Throw 'synthetic firewall query failure'
        }

        It 'rejects any enabled outbound allow-bypass rule before isolation' {
            Mock Get-NetFirewallRule {
                return @(
                    [pscustomobject]@{ Name = 'ordinary-allow' },
                    [pscustomobject]@{ Name = 'authenticated-bypass' }
                )
            }
            Mock Get-RSFirewallSecurityFilters {
                return [pscustomobject]@{
                    OverrideBlockRules = ($Rule.Name -eq 'authenticated-bypass')
                }
            }

            {
                Assert-RSNoOutboundBlockBypass
            } | Should Throw 'Enabled outbound authenticated-bypass rules can override a block rule: authenticated-bypass.'
        }

        It 'records a complete scan when every enabled outbound allow rule cannot override blocks' {
            Mock Get-NetFirewallRule {
                return @(
                    [pscustomobject]@{ Name = 'ordinary-allow-1' },
                    [pscustomobject]@{ Name = 'ordinary-allow-2' }
                )
            }
            Mock Get-RSFirewallSecurityFilters {
                return [pscustomobject]@{ OverrideBlockRules = $false }
            }

            $scan = Assert-RSNoOutboundBlockBypass

            $scan.enabledOutboundAllowRulesScanned | Should Be 2
            $scan.outboundBlockBypassRules | Should Be 0
            Assert-MockCalled Get-RSFirewallSecurityFilters -Times 2 -Scope It
        }
    }

    Context 'attestation and fail-closed behavior' {
        BeforeEach {
            Mock Test-RSAdministrator { return $true }
            Mock Assert-RSFirewallProviderAvailable {
                return [pscustomobject]@{
                    service = 'MpsSvc'
                    serviceStatus = 'Running'
                    policyStore = 'PersistentStore'
                    effectivePolicyStore = 'ActiveStore'
                    enabledProfiles = @('Domain', 'Private', 'Public')
                }
            }
            Mock New-RSOutboundBlockRule {}
            Mock Assert-RSOutboundBlockRule {}
            Mock Assert-RSNoOutboundBlockBypass {
                return [pscustomobject]@{
                    enabledOutboundAllowRulesScanned = 12
                    outboundBlockBypassRules = 0
                }
            }
            Mock Remove-RSExactFirewallRules {}

            $script:canaryCall = 0
            $script:canarySequence = @($true, $false, $true)
            Mock Invoke-RSExternalConnectCanary {
                $connected = $script:canarySequence[$script:canaryCall]
                $script:canaryCall++
                return [pscustomobject]@{
                    connected = $connected
                    exitCode = if ($connected) { 0 } else { 30 }
                    timedOut = $false
                    stdout = if ($connected) { 'connected' } else { '' }
                    stderr = if ($connected) { '' } else { 'System.Net.Sockets.SocketException' }
                    durationMilliseconds = '7'
                }
            }
            Mock Invoke-RSRootCommand {
                return [pscustomobject]@{
                    exitCode = 17
                    timedOut = $false
                    stdout = 'root-stdout'
                    stderr = 'root-stderr'
                    stdoutSha256 = ('a' * 64)
                    stderrSha256 = ('b' * 64)
                    startedUtc = '2026-07-30T00:00:00.0000000Z'
                    completedUtc = '2026-07-30T00:00:01.0000000Z'
                    durationMilliseconds = '1000'
                }
            }
        }

        It 'returns one complete runner-enforced attestation while preserving the root exit and output' {
            $attestation = Invoke-RSIsolatedProcess `
                -FilePath $script:pwshPath `
                -ArgumentList @('-NoProfile', '-Command', 'exit 17') `
                -CanaryExecutablePath $script:pwshPath `
                -WorkingDirectory $script:workingDirectory `
                -ExclusiveDisposableRunner

            @($attestation).Count | Should Be 1
            $attestation.formatId | Should Be 'red-salamander-runner-network-isolation-attestation'
            $attestation.provider | Should Be 'windows-defender-firewall-host-wide-outbound-block'
            $attestation.enforcement | Should Be 'runner-enforced'
            $attestation.providerStatus | Should Be 'pass'
            $attestation.isolationVerified | Should Be $true
            $attestation.rootCommandExecuted | Should Be $true
            $attestation.ruleSetupVerified | Should Be $true
            $attestation.ruleRemovalVerified | Should Be $true
            $attestation.programIdentityStable | Should Be $true
            @($attestation.programs).Count | Should Be 1
            @($attestation.programsAfter).Count | Should Be 1
            $attestation.programs[0].sha256 | Should Be $attestation.programsAfter[0].sha256
            $attestation.canary.beforeIsolation.connected | Should Be $true
            $attestation.canary.duringIsolation.connected | Should Be $false
            $attestation.canary.afterCleanup.connected | Should Be $true
            $attestation.rootCommand.result.exitCode | Should Be 17
            $attestation.rootCommand.result.stdout | Should Be 'root-stdout'
            $attestation.rootCommand.result.stderr | Should Be 'root-stderr'
            $attestation.policyBypassScan.beforeIsolation.outboundBlockBypassRules | Should Be 0
            $attestation.policyBypassScan.duringIsolation.outboundBlockBypassRules | Should Be 0
            $attestation.scopeContract.hostWide | Should Be $true
            $attestation.scopeContract.descendantCoverage | Should Be 'all-programs'
            $attestation.scopeContract.exactProgramPathsOnly | Should Be $false
            $attestation.scopeContract.completeExecutableClosureRequired | Should Be $false
            $attestation.scopeContract.unlistedExecutablesAreBlocked | Should Be $true
            $attestation.scopeContract.exclusiveDisposableRunnerAcknowledged | Should Be $true
            $attestation.scopeContract.maximumCommandDurationSeconds | Should Be 7200
            $attestation.scopeContract.proxyOrEnvironmentEvidenceAccepted | Should Be $false
            @($attestation.failures).Count | Should Be 0

            Assert-MockCalled New-RSOutboundBlockRule -Times 1 -Scope It
            Assert-MockCalled Assert-RSOutboundBlockRule -Times 1 -Scope It
            Assert-MockCalled Assert-RSNoOutboundBlockBypass -Times 2 -Scope It
            Assert-MockCalled Invoke-RSRootCommand -Times 1 -Scope It
            Assert-MockCalled Remove-RSExactFirewallRules -Times 1 -Scope It
            Assert-MockCalled Invoke-RSExternalConnectCanary -Times 3 -Scope It
        }

        It 'does not create rules or invoke the root when the before-isolation canary cannot connect' {
            $script:canarySequence = @($false)

            $attestation = Invoke-RSIsolatedProcess `
                -FilePath $script:pwshPath `
                -CanaryExecutablePath $script:pwshPath `
                -WorkingDirectory $script:workingDirectory `
                -ExclusiveDisposableRunner

            $attestation.providerStatus | Should Be 'fail'
            $attestation.isolationVerified | Should Be $false
            $attestation.rootCommandExecuted | Should Be $false
            @($attestation.failures).Count | Should Be 1
            $attestation.failures[0].stage | Should Be 'setup-canary-before'

            Assert-MockCalled New-RSOutboundBlockRule -Times 0 -Scope It
            Assert-MockCalled Invoke-RSRootCommand -Times 0 -Scope It
            Assert-MockCalled Remove-RSExactFirewallRules -Times 0 -Scope It
        }

        It 'removes its exact rules but never invokes the root when the blocked canary still connects' {
            $script:canarySequence = @($true, $true, $true)

            $attestation = Invoke-RSIsolatedProcess `
                -FilePath $script:pwshPath `
                -CanaryExecutablePath $script:pwshPath `
                -WorkingDirectory $script:workingDirectory `
                -ExclusiveDisposableRunner

            $attestation.providerStatus | Should Be 'fail'
            $attestation.isolationVerified | Should Be $false
            $attestation.rootCommandExecuted | Should Be $false
            $attestation.ruleRemovalVerified | Should Be $true
            $attestation.failures[0].stage | Should Be 'setup-canary-during'

            Assert-MockCalled Invoke-RSRootCommand -Times 0 -Scope It
            Assert-MockCalled Remove-RSExactFirewallRules -Times 1 -Scope It
        }

        It 'cleans the rule and never invokes the root if an outbound bypass appears during setup' {
            $script:bypassScanCall = 0
            Mock Assert-RSNoOutboundBlockBypass {
                $script:bypassScanCall++
                if ($script:bypassScanCall -eq 2) {
                    throw 'synthetic outbound bypass appeared'
                }
                return [pscustomobject]@{
                    enabledOutboundAllowRulesScanned = 12
                    outboundBlockBypassRules = 0
                }
            }
            $script:canarySequence = @($true, $true)

            $attestation = Invoke-RSIsolatedProcess `
                -FilePath $script:pwshPath `
                -CanaryExecutablePath $script:pwshPath `
                -WorkingDirectory $script:workingDirectory `
                -ExclusiveDisposableRunner

            $attestation.providerStatus | Should Be 'fail'
            $attestation.isolationVerified | Should Be $false
            $attestation.rootCommandExecuted | Should Be $false
            $attestation.ruleRemovalVerified | Should Be $true
            $attestation.failures[0].stage | Should Be 'firewall-setup'
            $attestation.failures[0].message | Should Match 'synthetic outbound bypass appeared'

            Assert-MockCalled Invoke-RSRootCommand -Times 0 -Scope It
            Assert-MockCalled Remove-RSExactFirewallRules -Times 1 -Scope It
            Assert-MockCalled Invoke-RSExternalConnectCanary -Times 2 -Scope It
        }

        It 'invalidates an otherwise completed run when exact-rule cleanup fails' {
            Mock Remove-RSExactFirewallRules { throw 'synthetic exact cleanup failure' }
            $script:canarySequence = @($true, $false)

            $attestation = Invoke-RSIsolatedProcess `
                -FilePath $script:pwshPath `
                -CanaryExecutablePath $script:pwshPath `
                -WorkingDirectory $script:workingDirectory `
                -ExclusiveDisposableRunner

            $attestation.providerStatus | Should Be 'fail'
            $attestation.isolationVerified | Should Be $false
            $attestation.rootCommandExecuted | Should Be $true
            $attestation.ruleRemovalVerified | Should Be $false
            $attestation.rootCommand.result.exitCode | Should Be 17
            @($attestation.failures).Count | Should Be 1
            $attestation.failures[0].stage | Should Be 'cleanup'
            $attestation.failures[0].message | Should Match 'synthetic exact cleanup failure'

            Assert-MockCalled Invoke-RSRootCommand -Times 1 -Scope It
            Assert-MockCalled Invoke-RSExternalConnectCanary -Times 2 -Scope It
        }
    }

    Context 'host-wide firewall-rule ownership' {
        BeforeEach {
            $script:testRule = [pscustomobject]@{
                name = 'RedSalamander.TerminalEngine.Gate0.0123456789abcdef0123456789abcdef.AllPrograms'
                displayName = 'RedSalamander Terminal Gate 0 whole-runner isolation'
                group = 'RedSalamander.TerminalEngine.Gate0.0123456789abcdef0123456789abcdef'
                description = 'Ephemeral all-program outbound block'
            }
        }

        It 'creates one outbound block for all programs and a unique group' {
            Mock Get-NetFirewallRule { return @() }
            Mock New-NetFirewallRule {}

            New-RSOutboundBlockRule -Rule $script:testRule

            Assert-MockCalled New-NetFirewallRule -Times 1 -Scope It -ParameterFilter {
                $PolicyStore -eq 'PersistentStore' -and
                $Name -eq $script:testRule.name -and
                $Group -eq $script:testRule.group -and
                $null -eq $Program -and
                [string]$Direction -eq 'Outbound' -and
                [string]$Action -eq 'Block' -and
                [string]$Enabled -eq 'True' -and
                [string]$Profile -eq 'Any' -and
                [string]$Protocol -eq 'Any'
            }
        }

        It 'verifies the effective rule covers every profile, interface, service, address, port, and protocol' {
            $activeRule = [pscustomobject]@{
                Name = $script:testRule.name
                Group = $script:testRule.group
                Enabled = 'True'
                Direction = 'Outbound'
                Action = 'Block'
                Profile = 'Any'
            }
            Mock Get-NetFirewallRule { return $activeRule }
            Mock Get-RSFirewallApplicationFilters {
                return [pscustomobject]@{ Program = 'Any'; Package = '' }
            }
            Mock Get-RSFirewallAddressFilters {
                return [pscustomobject]@{ LocalAddress = 'Any'; RemoteAddress = 'Any' }
            }
            Mock Get-RSFirewallPortFilters {
                return [pscustomobject]@{ Protocol = 'Any'; LocalPort = 'Any'; RemotePort = 'Any' }
            }
            Mock Get-RSFirewallServiceFilters {
                return [pscustomobject]@{ Service = 'Any' }
            }
            Mock Get-RSFirewallInterfaceFilters {
                return [pscustomobject]@{ InterfaceAlias = 'Any' }
            }
            Mock Get-RSFirewallInterfaceTypeFilters {
                return [pscustomobject]@{ InterfaceType = 'Any' }
            }
            Mock Get-RSFirewallSecurityFilters {
                return [pscustomobject]@{ OverrideBlockRules = 'False' }
            }

            Assert-RSOutboundBlockRule -Rule $script:testRule

            Assert-MockCalled Get-RSFirewallApplicationFilters -Times 1 -Scope It
            Assert-MockCalled Get-RSFirewallAddressFilters -Times 1 -Scope It
            Assert-MockCalled Get-RSFirewallPortFilters -Times 1 -Scope It
            Assert-MockCalled Get-RSFirewallServiceFilters -Times 1 -Scope It
            Assert-MockCalled Get-RSFirewallInterfaceFilters -Times 1 -Scope It
            Assert-MockCalled Get-RSFirewallInterfaceTypeFilters -Times 1 -Scope It
            Assert-MockCalled Get-RSFirewallSecurityFilters -Times 1 -Scope It
        }

        It 'removes only a rule whose generated name, group, and all-program scope still match' {
            $script:removed = $false
            $existingRule = [pscustomobject]@{
                Name = $script:testRule.name
                DisplayName = $script:testRule.displayName
                Group = $script:testRule.group
                Description = $script:testRule.description
                Direction = 'Outbound'
                Action = 'Block'
            }
            Mock Get-NetFirewallRule {
                if ($PolicyStore -eq 'PersistentStore' -and -not $script:removed) {
                    return $existingRule
                }
                return @()
            }
            Mock Get-RSFirewallApplicationFilters {
                return [pscustomobject]@{ Program = 'Any'; Package = '' }
            }
            Mock Remove-NetFirewallRule {
                $script:removed = $true
            }

            Remove-RSExactFirewallRules -Rules @($script:testRule)

            $script:removed | Should Be $true
            Assert-MockCalled Remove-NetFirewallRule -Times 1 -Scope It -ParameterFilter {
                $PolicyStore -eq 'PersistentStore' -and
                $Name -eq $script:testRule.name
            }
        }

        It 'refuses to remove a generated-name collision whose ownership identity changed' {
            $script:persistentLookup = 0
            $mutatedRule = [pscustomobject]@{
                Name = $script:testRule.name
                DisplayName = $script:testRule.displayName
                Group = 'Some.Other.Owner'
                Description = $script:testRule.description
                Direction = 'Outbound'
                Action = 'Block'
            }
            Mock Get-NetFirewallRule {
                if ($PolicyStore -eq 'PersistentStore' -and $script:persistentLookup -eq 0) {
                    $script:persistentLookup++
                    return $mutatedRule
                }
                return @()
            }
            Mock Get-RSFirewallApplicationFilters {
                return [pscustomobject]@{ Program = 'Any'; Package = '' }
            }
            Mock Remove-NetFirewallRule {}

            {
                Remove-RSExactFirewallRules -Rules @($script:testRule)
            } | Should Throw "Exact firewall-rule cleanup failed: $($script:testRule.name): The generated rule identity changed; refusing a broad or ambiguous removal."

            Assert-MockCalled Remove-NetFirewallRule -Times 0 -Scope It
        }
    }
}

if ([string]::Equals(
        [System.Environment]::GetEnvironmentVariable('REDSALAMANDER_RUN_FIREWALL_INTEGRATION'),
        '1',
        [System.StringComparison]::Ordinal)) {
    Describe 'Terminal-engine live firewall integration' -Tag RequiresElevation, RequiresExternalNetwork {
        It 'proves allowed-blocked-restored connectivity for an unlisted dynamic descendant' {
            . $helperScript -LoadOnly
            if (-not (Test-RSAdministrator)) {
                throw 'REDSALAMANDER_RUN_FIREWALL_INTEGRATION=1 requires an elevated test process.'
            }

            $pwshPath = [System.IO.Path]::GetFullPath(
                [System.Diagnostics.Process]::GetCurrentProcess().MainModule.FileName)
            $address = [System.Environment]::GetEnvironmentVariable('REDSALAMANDER_FIREWALL_CANARY_ADDRESS')
            if ([string]::IsNullOrWhiteSpace($address)) {
                $address = '1.1.1.1'
            }
            $publicAddress = Resolve-RSCanaryAddress -Address $address
            $windowsPowerShellPath = [System.IO.Path]::GetFullPath(
                "$env:SystemRoot\System32\WindowsPowerShell\v1.0\powershell.exe")
            $descendantBefore = Invoke-RSExternalConnectCanary `
                -CanonicalCanaryExecutablePath $windowsPowerShellPath `
                -Address $publicAddress `
                -Port 443 `
                -TimeoutMilliseconds 5000 `
                -CanonicalWorkingDirectory $repoRoot
            $descendantBefore.connected | Should Be $true

            $descendantCanary = @"
`$client = [System.Net.Sockets.TcpClient]::new([System.Net.Sockets.AddressFamily]::InterNetwork)
try {
    `$pending = `$client.BeginConnect('$address', 443, `$null, `$null)
    if (-not `$pending.AsyncWaitHandle.WaitOne(5000)) { exit 42 }
    `$client.EndConnect(`$pending)
    if (`$client.Connected) { exit 41 }
    exit 42
}
catch { exit 42 }
finally { `$client.Dispose() }
"@
            $encodedDescendantCanary = [Convert]::ToBase64String(
                [Text.Encoding]::Unicode.GetBytes($descendantCanary))
            $rootScript = @'
& '__CHILD_PATH__' -NoLogo -NoProfile -NonInteractive -EncodedCommand '__ENCODED_COMMAND__'
if ($LASTEXITCODE -ne 42) {
    [Console]::Error.Write("dynamic descendant was not blocked; exit=$LASTEXITCODE")
    exit 43
}
[Console]::Out.Write('isolated-root:descendant-blocked')
'@.Replace('__CHILD_PATH__', $windowsPowerShellPath).
                Replace('__ENCODED_COMMAND__', $encodedDescendantCanary)

            $attestation = Invoke-RSIsolatedProcess `
                -FilePath $pwshPath `
                -ArgumentList @('-NoProfile', '-NonInteractive', '-Command', $rootScript) `
                -CanaryExecutablePath $pwshPath `
                -WorkingDirectory $repoRoot `
                -CanaryAddress $address `
                -ExclusiveDisposableRunner

            $attestation.providerStatus | Should Be 'pass'
            $attestation.rootCommand.result.exitCode | Should Be 0
            $attestation.rootCommand.result.stdout | Should Be 'isolated-root:descendant-blocked'
            $attestation.ruleRemovalVerified | Should Be $true
            $attestation.scopeContract.descendantCoverage | Should Be 'all-programs'

            $descendantAfter = Invoke-RSExternalConnectCanary `
                -CanonicalCanaryExecutablePath $windowsPowerShellPath `
                -Address $publicAddress `
                -Port 443 `
                -TimeoutMilliseconds 5000 `
                -CanonicalWorkingDirectory $repoRoot
            $descendantAfter.connected | Should Be $true
        }
    }
}
