[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [ValidateSet('x64', 'ARM64')]
    [string]$Platform,

    [Parameter(Mandatory = $true)]
    [ValidateSet('Release', 'ASan Debug')]
    [string]$Configuration,

    [Parameter(Mandatory = $true)]
    [ValidateSet('functional', 'full')]
    [string]$MeasurementMode,

    [Parameter(Mandatory = $true)]
    [string]$MSBuildPath,

    [Parameter(Mandatory = $true)]
    [string]$HostHeaderIncludeRoot,

    [Parameter(Mandatory = $true)]
    [ValidatePattern('^[0-9a-f]{64}$')]
    [string]$ExpectedHostHeaderTreeSha256,

    [Parameter(Mandatory = $true)]
    [string]$FixtureIncludeRoot,

    [Parameter(Mandatory = $true)]
    [string]$AdapterDll,

    [Parameter(Mandatory = $true)]
    [string]$OutputDirectory,

    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string]$ExpectedCandidateId,

    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string]$ExpectedUpstreamPin,

    [Parameter(Mandatory = $true)]
    [ValidatePattern('^[0-9a-f]{64}$')]
    [string]$ExpectedSourceSha256,

    [Parameter(Mandatory = $true)]
    [ValidatePattern('^[0-9a-f]{64}$')]
    [string]$ExpectedAdapterHeaderSha256,

    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string]$ExpectedBuildIdentity,

    [Parameter(Mandatory = $true)]
    [ValidateRange(1, [uint32]::MaxValue)]
    [uint32]$ExpectedAdapterAbi,

    [Parameter(Mandatory = $true)]
    [ValidatePattern('^(?:0x)?[0-9a-fA-F]{1,16}$')]
    [string]$ExpectedCapabilities,

    [Parameter(Mandatory = $true)]
    [ValidatePattern('^[0-9a-f]{64}$')]
    [string]$ExpectedFixtureManifestSha256,

    [uint64]$AddedReleasePackageBytes,

    [uint64]$CrossedAbiElements,

    [uint64]$NonSystemRuntimeDependencies,

    [ValidateRange(1, 7200)]
    [int]$TimeoutSeconds = 3600
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
    $stdout = $stdoutTask.GetAwaiter().GetResult()
    $stderr = $stderrTask.GetAwaiter().GetResult()
    return [pscustomobject]@{
        ExitCode = $process.ExitCode
        Stdout = $stdout
        Stderr = $stderr
    }
}

$msbuild = Resolve-RequiredFile $MSBuildPath 'MSBuild'
Import-Module (
    Join-Path $PSScriptRoot 'TerminalEngineLifecycle.psm1')
$hostHeaders = Resolve-RSTerminalGate0HostHeaderRoot `
    -HostHeaderIncludeRoot $HostHeaderIncludeRoot `
    -ExpectedTreeSha256 $ExpectedHostHeaderTreeSha256
$fixtures = Resolve-RequiredDirectory `
    $FixtureIncludeRoot `
    'fixture include root'
$adapter = Resolve-RequiredFile $AdapterDll 'candidate adapter DLL'
$project = Resolve-RequiredFile `
    (Join-Path $PSScriptRoot 'TerminalEngineGate0Harness.vcxproj') `
    'neutral Gate-0 harness project'
$adapterHeader = Resolve-RequiredFile `
    (Join-Path $PSScriptRoot 'TerminalEngineGate0Adapter.h') `
    'neutral Gate-0 adapter header'
$fixtureHeader = Resolve-RequiredFile `
    (Join-Path $fixtures 'TerminalEngineFixtures.v1.h') `
    'generated fixture header'

$actualAdapterHeaderSha256 = (
    Get-FileHash -LiteralPath $adapterHeader -Algorithm SHA256
).Hash.ToLowerInvariant()
if ($actualAdapterHeaderSha256 -cne $ExpectedAdapterHeaderSha256) {
    throw (
        "Adapter header identity mismatch: expected '{0}', actual '{1}'." -f
        $ExpectedAdapterHeaderSha256,
        $actualAdapterHeaderSha256)
}

$null = $fixtureHeader
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
    "/p:FixtureIncludeRoot=$fixtures",
    "/p:HarnessOutputDirectory=$binaryDirectory",
    "/p:HarnessIntermediateDirectory=$intermediateDirectory"
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
        "Neutral Gate-0 harness build failed.`n{0}`n{1}" -f
        $build.Stdout,
        $build.Stderr)
}

$harness = Resolve-RequiredFile `
    (Join-Path $binaryDirectory 'TerminalEngineGate0Harness.exe') `
    'neutral Gate-0 harness executable'
$runArguments = @(
    '--adapter-dll', $adapter,
    '--expected-candidate-id', $ExpectedCandidateId,
    '--expected-upstream-pin', $ExpectedUpstreamPin,
    '--expected-source-sha256', $ExpectedSourceSha256,
    '--expected-adapter-header-sha256', $ExpectedAdapterHeaderSha256,
    '--expected-build-identity', $ExpectedBuildIdentity,
    '--expected-adapter-abi', $ExpectedAdapterAbi.ToString(
        [Globalization.CultureInfo]::InvariantCulture),
    '--expected-capabilities', $ExpectedCapabilities,
    '--expected-fixture-manifest-sha256',
        $ExpectedFixtureManifestSha256,
    '--measurement-mode', $MeasurementMode
)
if ($MeasurementMode -ceq 'full') {
    $runArguments += @(
        '--added-release-package-bytes',
            $AddedReleasePackageBytes.ToString(
                [Globalization.CultureInfo]::InvariantCulture),
        '--crossed-abi-elements',
            $CrossedAbiElements.ToString(
                [Globalization.CultureInfo]::InvariantCulture),
        '--non-system-runtime-dependencies',
            $NonSystemRuntimeDependencies.ToString(
                [Globalization.CultureInfo]::InvariantCulture)
    )
}
$run = Invoke-CapturedProcess `
    -FilePath $harness `
    -ArgumentList $runArguments `
    -TimeoutMilliseconds ($TimeoutSeconds * 1000)
if ([string]::IsNullOrWhiteSpace($run.Stdout)) {
    throw "Neutral Gate-0 harness emitted no JSON. $($run.Stderr)"
}
try {
    $result = $run.Stdout | ConvertFrom-Json -Depth 32
}
catch {
    throw "Neutral Gate-0 harness emitted invalid JSON. $($run.Stdout)"
}
if ($result.formatId -cne 'red-salamander-terminal-engine-gate0-harness' -or
    [int]$result.version -ne 1) {
    throw 'Neutral Gate-0 harness output format identity is invalid.'
}
if ($result.fixtureManifestSha256 -cne
    $ExpectedFixtureManifestSha256) {
    throw 'Neutral Gate-0 harness output names the wrong fixture manifest.'
}
if ($result.measurementMode -cne $MeasurementMode) {
    throw 'Neutral Gate-0 harness output names the wrong measurement mode.'
}

$rows = @($result.fixtureRows)
if ($run.ExitCode -eq 0) {
    if ($result.status -cne 'pass' -or $rows.Count -ne 43) {
        throw 'A successful neutral Gate-0 run must contain 43 passing rows.'
    }
    $rowKeys = @{}
    foreach ($row in $rows) {
        $key = '{0}/{1}' -f $row.fixtureId, $row.scheduleId
        if ($rowKeys.ContainsKey($key)) {
            throw "Neutral Gate-0 output contains duplicate row '$key'."
        }
        $rowKeys[$key] = $true
        if ($row.status -cne 'pass' -or
            $row.expectedStateSha256 -cne $row.actualStateSha256 -or
            $row.expectedResourceTraceSha256 -cne
                $row.actualResourceTraceSha256) {
            throw "Neutral Gate-0 row '$key' did not pass exactly."
        }
    }
    if (@($result.limitations).Count -ne 0) {
        throw 'A successful neutral Gate-0 run cannot retain limitations.'
    }
    $measurements = @($result.measurements)
    if ($MeasurementMode -ceq 'full') {
        $requiredMetrics = @(
            [pscustomobject]@{
                Id = 'create-destroy-us'
                Unit = 'microseconds'
                Warmups = 10
                Samples = 50
            },
            [pscustomobject]@{
                Id = 'parse-8mib-us'
                Unit = 'microseconds'
                Warmups = 5
                Samples = 30
            },
            [pscustomobject]@{
                Id = 'full-snapshot-us'
                Unit = 'microseconds'
                Warmups = 10
                Samples = 50
            },
            [pscustomobject]@{
                Id = 'dirty-snapshot-us'
                Unit = 'microseconds'
                Warmups = 10
                Samples = 50
            },
            [pscustomobject]@{
                Id = 'resize-reflow-us'
                Unit = 'microseconds'
                Warmups = 5
                Samples = 30
            },
            [pscustomobject]@{
                Id = 'kitty-1024-rgba-us'
                Unit = 'microseconds'
                Warmups = 3
                Samples = 20
            },
            [pscustomobject]@{
                Id = 'eight-session-storm-us'
                Unit = 'microseconds'
                Warmups = 3
                Samples = 20
            },
            [pscustomobject]@{
                Id = 'eight-session-retained-bytes'
                Unit = 'bytes'
                Warmups = 1
                Samples = 5
            },
            [pscustomobject]@{
                Id = 'added-release-package-bytes'
                Unit = 'bytes'
                Warmups = 0
                Samples = 1
            },
            [pscustomobject]@{
                Id = 'crossed-abi-elements'
                Unit = 'count'
                Warmups = 0
                Samples = 1
            },
            [pscustomobject]@{
                Id = 'non-system-runtime-dependencies'
                Unit = 'count'
                Warmups = 0
                Samples = 1
            }
        )
        if ($measurements.Count -ne $requiredMetrics.Count -or
            [uint64]$result.qpcFrequency -eq 0) {
            throw 'A full neutral Gate-0 run must contain 11 QPC measurements.'
        }
        for ($index = 0; $index -lt $requiredMetrics.Count; $index++) {
            $expected = $requiredMetrics[$index]
            $actual = $measurements[$index]
            $samples = @($actual.samples)
            if ($actual.metricId -cne $expected.Id -or
                $actual.unit -cne $expected.Unit -or
                [int]$actual.warmupCount -ne $expected.Warmups -or
                $samples.Count -ne $expected.Samples) {
                throw "Neutral Gate-0 measurement '$index' is invalid."
            }
            foreach ($sample in $samples) {
                $parsed = [uint64]0
                if ($sample -isnot [string] -or
                    $sample -cnotmatch '^(?:0|[1-9][0-9]{0,19})$' -or
                    -not [uint64]::TryParse(
                        $sample,
                        [Globalization.NumberStyles]::None,
                        [Globalization.CultureInfo]::InvariantCulture,
                        [ref]$parsed)) {
                    throw (
                        "Neutral Gate-0 measurement '$index' contains " +
                        'a non-canonical uint64 sample.')
                }
            }
        }
    }
    elseif ($measurements.Count -ne 0) {
        throw 'A functional neutral Gate-0 run cannot contain measurements.'
    }
}
elseif ($result.status -cne 'fail') {
    throw 'A failing neutral Gate-0 process must report status=fail.'
}

$resultPath = Join-Path $output 'terminal-engine-gate0-result.json'
[IO.File]::WriteAllText(
    $resultPath,
    $run.Stdout,
    [Text.UTF8Encoding]::new($false))
if (-not [string]::IsNullOrWhiteSpace($run.Stderr)) {
    [Console]::Error.Write($run.Stderr)
}
Write-Output $resultPath
exit $run.ExitCode
