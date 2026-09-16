[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [ValidateSet('x64', 'ARM64')]
    [string]$Platform,

    [Parameter(Mandatory = $true)]
    [ValidateSet('Debug', 'Release', 'ASan Debug')]
    [string]$Configuration,

    [Parameter(Mandatory = $true)]
    [ValidateSet('layout', 'fuzz', 'c7')]
    [string]$Mode,

    [Parameter(Mandatory = $true)]
    [string]$MSBuildPath,

    [Parameter(Mandatory = $true)]
    [string]$HostHeaderIncludeRoot,

    [Parameter(Mandatory = $true)]
    [ValidatePattern('^[0-9a-f]{64}$')]
    [string]$ExpectedHostHeaderTreeSha256,

    [Parameter(Mandatory = $true)]
    [string]$OutputDirectory,

    [string]$AdapterDll,

    [ValidatePattern('^layout-v1-[0-9a-f]{64}$')]
    [string]$ExpectedLayoutIdentity,

    [ValidatePattern('^[\x21-\x7e]{1,256}$')]
    [string]$ExpectedCandidateId,

    [ValidatePattern('^[\x21-\x7e]{1,256}$')]
    [string]$ExpectedUpstreamPin,

    [ValidatePattern('^[0-9a-f]{64}$')]
    [string]$ExpectedSourceSha256,

    [ValidatePattern('^[0-9a-f]{64}$')]
    [string]$ExpectedAdapterHeaderSha256,

    [ValidatePattern('^build-v1-[0-9a-f]{64}$')]
    [string]$ExpectedBuildIdentity,

    [ValidatePattern('^(?:0x[0-9a-fA-F]{1,16}|[0-9]{1,20})$')]
    [string]$ExpectedCapabilities,

    [ValidatePattern('^(?:0x)?[0-9a-fA-F]{1,16}$')]
    [string]$Seed = '0x5253474154453031',

    [ValidateRange(1, 64)]
    [int]$RandomCases = 24,

    [ValidateRange(1, 1024)]
    [int]$MaximumCaseBytes = 256,

    [ValidateRange(1, 1800)]
    [int]$TimeoutSeconds = 600
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

function Resolve-RequiredFile {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path,

        [Parameter(Mandatory = $true)]
        [string]$Description
    )

    $resolved = [IO.Path]::GetFullPath($Path)
    if (-not [IO.File]::Exists($resolved)) {
        throw [IO.FileNotFoundException]::new(
            "$Description does not exist.",
            $resolved)
    }
    return $resolved
}

function Resolve-RequiredDirectory {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path,

        [Parameter(Mandatory = $true)]
        [string]$Description
    )

    $resolved = [IO.Path]::GetFullPath($Path)
    if (-not [IO.Directory]::Exists($resolved)) {
        throw [IO.DirectoryNotFoundException]::new(
            "$Description does not exist.")
    }
    return $resolved
}

function Invoke-CapturedProcess {
    param(
        [Parameter(Mandatory = $true)]
        [string]$FilePath,

        [Parameter(Mandatory = $true)]
        [string[]]$ArgumentList,

        [Parameter(Mandatory = $true)]
        [int]$TimeoutMilliseconds
    )

    $startInfo = [Diagnostics.ProcessStartInfo]::new()
    $startInfo.FileName = $FilePath
    $startInfo.UseShellExecute = $false
    $startInfo.CreateNoWindow = $true
    $startInfo.RedirectStandardOutput = $true
    $startInfo.RedirectStandardError = $true
    foreach ($argument in $ArgumentList) {
        [void]$startInfo.ArgumentList.Add($argument)
    }

    $process = [Diagnostics.Process]::new()
    $process.StartInfo = $startInfo
    if (-not $process.Start()) {
        throw "Failed to start '$FilePath'."
    }
    $stdoutTask = $process.StandardOutput.ReadToEndAsync()
    $stderrTask = $process.StandardError.ReadToEndAsync()
    if (-not $process.WaitForExit($TimeoutMilliseconds)) {
        $process.Kill($true)
        $process.WaitForExit()
        throw "'$FilePath' exceeded the timeout."
    }
    return [pscustomobject]@{
        ExitCode = $process.ExitCode
        Stdout = $stdoutTask.GetAwaiter().GetResult()
        Stderr = $stderrTask.GetAwaiter().GetResult()
    }
}

function Get-EmbeddedLayoutJcs {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Json
    )

    $prefix = ',"layoutJcs":'
    $suffix = ',"layoutSha256":'
    $start = $Json.IndexOf(
        $prefix,
        [StringComparison]::Ordinal)
    if ($start -lt 0) {
        throw 'Layout output does not contain layoutJcs.'
    }
    $start += $prefix.Length
    $end = $Json.IndexOf(
        $suffix,
        $start,
        [StringComparison]::Ordinal)
    if ($end -le $start) {
        throw 'Layout output does not terminate layoutJcs canonically.'
    }
    return $Json.Substring($start, $end - $start)
}

function Get-Sha256Hex {
    param(
        [Parameter(Mandatory = $true)]
        [byte[]]$Bytes
    )

    return [Convert]::ToHexString(
        [Security.Cryptography.SHA256]::HashData($Bytes)
    ).ToLowerInvariant()
}

$msbuild = Resolve-RequiredFile $MSBuildPath 'MSBuild'
Import-Module (
    Join-Path $PSScriptRoot 'TerminalEngineLifecycle.psm1')
$hostHeaders = Resolve-RSTerminalGate0HostHeaderRoot `
    -HostHeaderIncludeRoot $HostHeaderIncludeRoot `
    -ExpectedTreeSha256 $ExpectedHostHeaderTreeSha256
$project = Resolve-RequiredFile `
    (Join-Path $PSScriptRoot 'TerminalEngineGate0Probe.vcxproj') `
    'Gate-0 probe project'
$output = [IO.Path]::GetFullPath($OutputDirectory)
$binaryDirectory = Join-Path $output 'bin'
$intermediateDirectory = Join-Path $output 'obj'
[void][IO.Directory]::CreateDirectory($binaryDirectory)
[void][IO.Directory]::CreateDirectory($intermediateDirectory)

$buildArguments = @(
    $project,
    '/m',
    '/nologo',
    '/t:Build',
    '/v:minimal',
    "/p:Configuration=$Configuration;Platform=$Platform",
    '/p:RSUseVcpkgManifest=false',
    '/p:VcpkgEnabled=false',
    '/p:VcpkgEnableManifest=false',
    "/p:HostHeaderIncludeRoot=$($hostHeaders.IncludeRoot)",
    "/p:ProbeOutputDirectory=$binaryDirectory",
    "/p:ProbeIntermediateDirectory=$intermediateDirectory"
)
$null = Resolve-RSTerminalGate0HostHeaderRoot `
    -HostHeaderIncludeRoot $hostHeaders.IncludeRoot `
    -ExpectedTreeSha256 $ExpectedHostHeaderTreeSha256
$build = Invoke-CapturedProcess `
    -FilePath $msbuild `
    -ArgumentList $buildArguments `
    -TimeoutMilliseconds ($TimeoutSeconds * 1000)
$null = Resolve-RSTerminalGate0HostHeaderRoot `
    -HostHeaderIncludeRoot $hostHeaders.IncludeRoot `
    -ExpectedTreeSha256 $ExpectedHostHeaderTreeSha256
if ($build.ExitCode -ne 0) {
    throw (
        "Gate-0 probe build failed.`n{0}`n{1}" -f
        $build.Stdout,
        $build.Stderr)
}

$probe = Resolve-RequiredFile `
    (Join-Path $binaryDirectory 'TerminalEngineGate0Probe.exe') `
    'Gate-0 probe executable'
$runArguments = @('--mode', $Mode)
$expectedC7StreamSha256 = ''
if ($Mode -cin @('fuzz', 'c7')) {
    if ([string]::IsNullOrWhiteSpace($AdapterDll) -or
        [string]::IsNullOrWhiteSpace($ExpectedLayoutIdentity) -or
        [string]::IsNullOrWhiteSpace($ExpectedCandidateId) -or
        [string]::IsNullOrWhiteSpace($ExpectedUpstreamPin) -or
        [string]::IsNullOrWhiteSpace($ExpectedSourceSha256) -or
        [string]::IsNullOrWhiteSpace(
            $ExpectedAdapterHeaderSha256) -or
        [string]::IsNullOrWhiteSpace($ExpectedBuildIdentity) -or
        [string]::IsNullOrWhiteSpace($ExpectedCapabilities)) {
        throw (
            'Adapter mode requires the adapter, layout, candidate, pin, ' +
            'source, header, build, and capability identities.')
    }
    $adapter = Resolve-RequiredFile `
        $AdapterDll `
        'Gate-0 adapter DLL'
    $runArguments += @(
        '--adapter-dll', $adapter,
        '--expected-layout-identity', $ExpectedLayoutIdentity,
        '--expected-candidate-id', $ExpectedCandidateId,
        '--expected-upstream-pin', $ExpectedUpstreamPin,
        '--expected-source-sha256', $ExpectedSourceSha256,
        '--expected-adapter-header-sha256',
            $ExpectedAdapterHeaderSha256,
        '--expected-build-identity', $ExpectedBuildIdentity,
        '--expected-architecture', $Platform,
        '--expected-capabilities', $ExpectedCapabilities
    )
    if ($Mode -ceq 'fuzz') {
        $runArguments += @(
            '--seed', $Seed,
            '--random-cases',
                $RandomCases.ToString(
                    [Globalization.CultureInfo]::InvariantCulture),
            '--maximum-case-bytes',
                $MaximumCaseBytes.ToString(
                    [Globalization.CultureInfo]::InvariantCulture)
        )
    }
    else {
        Import-Module (
            Join-Path $PSScriptRoot 'TerminalEngineC7Proof.psm1') -Force
        $fixedStream = Get-RSTerminalEngineC7FixedStreamBytes `
            -RepoRoot ([IO.Path]::GetFullPath((
                Join-Path $PSScriptRoot '..\..')))
        $expectedC7StreamSha256 = Get-Sha256Hex -Bytes $fixedStream
        $runArguments += @(
            '--c7-stream-hex',
            [Convert]::ToHexString($fixedStream).ToLowerInvariant()
        )
    }
}
elseif (-not [string]::IsNullOrWhiteSpace($AdapterDll) -or
    -not [string]::IsNullOrWhiteSpace($ExpectedLayoutIdentity) -or
    -not [string]::IsNullOrWhiteSpace($ExpectedCandidateId) -or
    -not [string]::IsNullOrWhiteSpace($ExpectedUpstreamPin) -or
    -not [string]::IsNullOrWhiteSpace($ExpectedSourceSha256) -or
    -not [string]::IsNullOrWhiteSpace(
        $ExpectedAdapterHeaderSha256) -or
    -not [string]::IsNullOrWhiteSpace($ExpectedBuildIdentity) -or
    -not [string]::IsNullOrWhiteSpace($ExpectedCapabilities)) {
    throw 'Layout mode does not accept adapter identity inputs.'
}

$run = Invoke-CapturedProcess `
    -FilePath $probe `
    -ArgumentList $runArguments `
    -TimeoutMilliseconds ($TimeoutSeconds * 1000)
if ([string]::IsNullOrWhiteSpace($run.Stdout)) {
    throw "Gate-0 probe emitted no JSON. $($run.Stderr)"
}
try {
    $result = $run.Stdout | ConvertFrom-Json -Depth 64
}
catch {
    throw "Gate-0 probe emitted invalid JSON. $($run.Stdout)"
}

if ($Mode -ceq 'layout') {
    if ($result.formatId -cne
            'red-salamander-terminal-engine-gate0-layout' -or
        [int]$result.version -ne 1 -or
        $result.status -cne 'pass' -or
        $result.layoutIdentity -cnotmatch
            '^layout-v1-[0-9a-f]{64}$' -or
        $result.layoutSha256 -cnotmatch '^[0-9a-f]{64}$') {
        throw 'Gate-0 layout output identity is invalid.'
    }
    $layoutJcs = Get-EmbeddedLayoutJcs -Json $run.Stdout.Trim()
    $actualSha256 = Get-Sha256Hex -Bytes (
        [Text.UTF8Encoding]::new($false).GetBytes($layoutJcs))
    if ($actualSha256 -cne $result.layoutSha256 -or
        $result.layoutIdentity -cne
            "layout-v1-$actualSha256") {
        throw 'Gate-0 layout JCS digest does not match its identity.'
    }
    if (@($result.layoutJcs.enums).Count -ne 6 -or
        @($result.layoutJcs.functionPointers).Count -ne 9 -or
        @($result.layoutJcs.structs).Count -ne 8) {
        throw 'Gate-0 layout output does not cover the complete public ABI.'
    }
    $resultLeaf = 'terminal-engine-gate0-layout.json'
}
elseif ($Mode -ceq 'fuzz') {
    $expectedCapabilitiesValue = if (
        $ExpectedCapabilities.StartsWith(
            '0x',
            [StringComparison]::OrdinalIgnoreCase)) {
        [Convert]::ToUInt64(
            $ExpectedCapabilities.Substring(2),
            16)
    }
    else {
        [uint64]::Parse(
            $ExpectedCapabilities,
            [Globalization.NumberStyles]::None,
            [Globalization.CultureInfo]::InvariantCulture)
    }
    if ($result.formatId -cne
            'red-salamander-terminal-engine-gate0-fuzz' -or
        [int]$result.version -ne 1 -or
        $result.layoutIdentity -cne $ExpectedLayoutIdentity -or
        $result.adapterLayoutSha256 -cne
            $ExpectedLayoutIdentity.Substring('layout-v1-'.Length) -or
        $result.candidateId -cne $ExpectedCandidateId -or
        $result.upstreamPin -cne $ExpectedUpstreamPin -or
        $result.sourceSha256 -cne $ExpectedSourceSha256 -or
        $result.adapterHeaderSha256 -cne
            $ExpectedAdapterHeaderSha256 -or
        $result.buildIdentity -cne $ExpectedBuildIdentity -or
        $result.architecture -cne $Platform -or
        [uint64]$result.capabilities -ne
            $expectedCapabilitiesValue -or
        $result.fixedStreamSha256 -cne
            $expectedC7StreamSha256 -or
        [int]$result.repeatCount -ne 2 -or
        [int]$result.caseCount -lt 6 -or
        [uint64]$result.denialSweepCount -eq 0 -or
        [uint64]$result.runCount -ne
            [uint64]$result.unloadCount -or
        $result.adapterIdentitySha256 -cnotmatch
            '^[0-9a-f]{64}$' -or
        $result.adapterLayoutSha256 -cnotmatch
            '^[0-9a-f]{64}$' -or
        $result.resourceTraceSha256 -cnotmatch
            '^[0-9a-f]{64}$' -or
        $result.resultTraceSha256 -cnotmatch
            '^[0-9a-f]{64}$' -or
        $result.stateTraceSha256 -cnotmatch
            '^[0-9a-f]{64}$') {
        throw 'Gate-0 fuzz output contract is invalid.'
    }
    $resultLeaf = 'terminal-engine-gate0-fuzz.json'
}
else {
    $expectedCapabilitiesValue = if (
        $ExpectedCapabilities.StartsWith(
            '0x',
            [StringComparison]::OrdinalIgnoreCase)) {
        [Convert]::ToUInt64(
            $ExpectedCapabilities.Substring(2),
            16)
    }
    else {
        [uint64]::Parse(
            $ExpectedCapabilities,
            [Globalization.NumberStyles]::None,
            [Globalization.CultureInfo]::InvariantCulture)
    }
    $expectedSchedules = [string[]]@(
        'all-at-once',
        'bytewise',
        'terminator-split')
    $scheduleStates = @($result.scheduleStates)
    $actualSchedules = [string[]]@(
        $scheduleStates | ForEach-Object { [string]$_.scheduleId })
    $stateSizesValid = $true
    $strictUtf8 = [Text.UTF8Encoding]::new($false, $true)
    foreach ($state in $scheduleStates) {
        if ($state.stateJcs -isnot [string] -or
            $strictUtf8.GetByteCount([string]$state.stateJcs) -gt 8192) {
            $stateSizesValid = $false
            break
        }
    }
    if ($result.formatId -cne
            'red-salamander-terminal-engine-gate0-c7' -or
        [int]$result.version -ne 1 -or
        $result.candidateId -cne $ExpectedCandidateId -or
        $result.upstreamPin -cne $ExpectedUpstreamPin -or
        $result.sourceSha256 -cne $ExpectedSourceSha256 -or
        $result.adapterHeaderSha256 -cne
            $ExpectedAdapterHeaderSha256 -or
        $result.buildIdentity -cne $ExpectedBuildIdentity -or
        $result.architecture -cne $Platform -or
        [uint64]$result.capabilities -ne
            $expectedCapabilitiesValue -or
        $result.adapterLayoutSha256 -cne
            $ExpectedLayoutIdentity.Substring('layout-v1-'.Length) -or
        $result.adapterIdentitySha256 -cnotmatch
            '^[0-9a-f]{64}$' -or
        [int]$result.sessionCount -ne 4 -or
        [int]$result.quietCount -ne 4 -or
        [int]$result.unloadCount -ne 1 -or
        ($actualSchedules -join "`0") -cne
            ($expectedSchedules -join "`0") -or
        -not $stateSizesValid -or
        $result.orderingStateJcs -isnot [string] -or
        $strictUtf8.GetByteCount(
            [string]$result.orderingStateJcs) -gt 8192 -or
        (@($result.orderingActionFlags) -join "`0") -cne
            (@(
                'action-observed',
                'checked-arithmetic',
                'order-preserved') -join "`0")) {
        throw 'Gate-0 C7 output contract is invalid.'
    }
    $resultLeaf = 'terminal-engine-gate0-c7.json'
}

$resultPath = Join-Path $output $resultLeaf
[IO.File]::WriteAllText(
    $resultPath,
    $run.Stdout,
    [Text.UTF8Encoding]::new($false))
if (-not [string]::IsNullOrWhiteSpace($run.Stderr)) {
    [Console]::Error.Write($run.Stderr)
}
if ($run.ExitCode -ne 0 -or $result.status -cne 'pass') {
    throw "Gate-0 $Mode probe failed with exit code $($run.ExitCode)."
}
Write-Output $resultPath
