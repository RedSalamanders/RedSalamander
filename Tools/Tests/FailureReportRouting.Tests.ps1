Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
Import-Module (Join-Path $PSScriptRoot '../Modules/Build/RedMonitorDependency.psm1') -Force
$routingCall = 'RedSalamander::TestSupport::RouteFailureReportsAwayFromDialogs();'
$selfTestSwitch = '--failure-report-self-test'
$selfTestTimeoutMilliseconds = 30000

function Get-RSRoutingSource {
    param(
        [Parameter(Mandatory = $true)]
        [string]$RelativePath
    )

    return [System.IO.File]::ReadAllText((Resolve-RSRedMonitorTestInput -RepoRoot $repoRoot -Path $RelativePath))
}

# The preprocessor conditions in force at the first line that contains $Text, and the lines just before it.
function Get-RSRoutingContext {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Source,

        [Parameter(Mandatory = $true)]
        [string]$Text
    )

    $lines = $Source -split "\r?\n"
    $conditions = [System.Collections.Generic.List[string]]::new()
    for ($index = 0; $index -lt $lines.Count; ++$index) {
        $line = $lines[$index].Trim()
        if ($line -match '^#\s*(?<keyword>if|ifdef|ifndef)\b(?<rest>.*)$') {
            $conditions.Add(('{0} {1}' -f $Matches['keyword'], $Matches['rest'].Trim()).Trim())
        }
        elseif ($line -match '^#\s*endif\b') {
            if ($conditions.Count -gt 0) {
                $conditions.RemoveAt($conditions.Count - 1)
            }
        }
        elseif ($line -match '^#\s*(else|elif)\b') {
            if ($conditions.Count -gt 0) {
                $conditions[$conditions.Count - 1] = 'else ' + $conditions[$conditions.Count - 1]
            }
        }
        elseif ($line.Contains($Text)) {
            $firstBefore = [Math]::Max(0, $index - 4)
            $before = if ($index -gt 0) { $lines[$firstBefore..($index - 1)] -join "`n" } else { '' }
            return [pscustomobject]@{
                Conditions = @($conditions)
                Before = $before
            }
        }
    }

    return $null
}

# Every native test project: an Application project owns one wmain/main, a DynamicLibrary (CppUnit) project one module initializer.
function Get-RSNativeTestProjects {
    $projects = [System.Collections.Generic.List[object]]::new()
    foreach ($directory in (Get-ChildItem -LiteralPath (Join-Path $repoRoot 'Tests') -Directory | Sort-Object Name)) {
        $projectPath = Join-Path $directory.FullName ($directory.Name + '.vcxproj')
        if (-not (Test-Path -LiteralPath $projectPath -PathType Leaf)) {
            continue
        }

        $projectText = [System.IO.File]::ReadAllText($projectPath)
        $kind = [regex]::Match($projectText, '<ConfigurationType>(?<kind>[A-Za-z]+)</ConfigurationType>').Groups['kind'].Value
        $entries = [System.Collections.Generic.List[object]]::new()
        $initializers = [System.Collections.Generic.List[object]]::new()
        foreach ($source in (Get-ChildItem -LiteralPath $directory.FullName -Filter '*.cpp' -File)) {
            $text = [System.IO.File]::ReadAllText($source.FullName)
            foreach ($match in [regex]::Matches($text, '(?m)^int\s+(?:__cdecl\s+)?w?main\s*\([^)]*\)\s*\r?\n\{[ \t]*\r?\n(?<first>[^\r\n]*)')) {
                $entries.Add([pscustomobject]@{ File = $source.Name; FirstStatement = $match.Groups['first'].Value.Trim() })
            }
            foreach ($match in [regex]::Matches($text, '(?m)^TEST_MODULE_INITIALIZE\(\w+\)\s*\r?\n\{[ \t]*\r?\n(?<first>[^\r\n]*)')) {
                $initializers.Add([pscustomobject]@{ File = $source.Name; FirstStatement = $match.Groups['first'].Value.Trim() })
            }
        }

        $projects.Add([pscustomobject]@{
                Name = $directory.Name
                Kind = $kind
                Entries = @($entries)
                Initializers = @($initializers)
            })
    }

    return @($projects)
}

function Test-RSExecutableContainsSelfTestSwitch {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    # The switch is a wide literal in the executable; an executable built before it existed cannot honor it. ISO-8859-1 maps every
    # byte to one character, so the UTF-16 bytes of the switch can be searched for as a string.
    $byteEncoding = [System.Text.Encoding]::GetEncoding(28591)
    $pattern = $byteEncoding.GetString([System.Text.Encoding]::Unicode.GetBytes($selfTestSwitch))
    return $byteEncoding.GetString([System.IO.File]::ReadAllBytes($Path)).Contains($pattern)
}

# Runs the executable with the self-test switch and ends only that exact process when the run waits instead of ending: a failed
# runtime check that opens a dialog never ends by itself.
function Invoke-RSBoundedFailureReportRun {
    param(
        [Parameter(Mandatory = $true)]
        [string]$ExecutablePath,

        [int]$TimeoutMilliseconds = $selfTestTimeoutMilliseconds
    )

    $startInfo = [System.Diagnostics.ProcessStartInfo]::new($ExecutablePath, $selfTestSwitch)
    $startInfo.UseShellExecute = $false
    $startInfo.RedirectStandardOutput = $true
    $startInfo.RedirectStandardError = $true
    $startInfo.CreateNoWindow = $true
    $startInfo.WorkingDirectory = Split-Path -Parent $ExecutablePath

    $process = [System.Diagnostics.Process]::new()
    $process.StartInfo = $startInfo
    $stopwatch = [System.Diagnostics.Stopwatch]::StartNew()
    $launchedAt = [DateTime]::Now
    try {
        [void]$process.Start()
        $standardOutput = $process.StandardOutput.ReadToEndAsync()
        $standardError = $process.StandardError.ReadToEndAsync()
        $exited = $process.WaitForExit($TimeoutMilliseconds)
        if ($exited) {
            $process.WaitForExit()
        }
        else {
            # Only the process object this function started may be ended, and only while its image and start time are still the
            # ones launched here; anything else is reported and left alone. No process is ever found by name or by walking a tree.
            $imagePath = $null
            try { $imagePath = $process.MainModule.FileName } catch [System.ComponentModel.Win32Exception] { $imagePath = $null } catch [System.InvalidOperationException] { $imagePath = $null }
            $isLaunchedImage = $null -ne $imagePath -and [string]::Equals($imagePath, [System.IO.Path]::GetFullPath($ExecutablePath), [System.StringComparison]::OrdinalIgnoreCase)
            $isLaunchedNow = $process.StartTime -ge $launchedAt.AddSeconds(-5) -and $process.StartTime -le [DateTime]::Now
            if (-not ($isLaunchedImage -and $isLaunchedNow)) {
                throw "The failure-report self-test (process $($process.Id)) did not end within $TimeoutMilliseconds ms and could not be confirmed as the process this test started; it was left running."
            }
            $process.Kill()
            [void]$process.WaitForExit(10000)
        }

        return [pscustomobject]@{
            Exited = $exited
            ExitCode = if ($exited) { $process.ExitCode } else { $null }
            Seconds = $stopwatch.Elapsed.TotalSeconds
            Output = (($standardOutput.GetAwaiter().GetResult()) + ($standardError.GetAwaiter().GetResult())) -replace '\s+', ' '
        }
    }
    finally {
        $process.Dispose()
    }
}

$nativeTestProjects = @(Get-RSNativeTestProjects)

Describe 'Failure-report routing keeps test processes off the desktop' {
    Context 'the routing header' {
        $header = Get-RSRoutingSource -RelativePath 'Tests\TestSupport\TestSupport.Tests.FailureReports.h'

        It 'includes what its functions use, so that it compiles in a translation unit that includes nothing before it' {
            # std::fputws is declared by <cwchar>, std::fflush and stderr by <cstdio>, std::abort by <cstdlib>.
            foreach ($include in @('<crtdbg.h>', '<cstdio>', '<cstdlib>', '<cwchar>', '<windows.h>')) {
                $header | Should Match ('(?m)^#include ' + [regex]::Escape($include) + '\s*$')
            }
        }

        It 'clears the CRT abort message and fault report, and keeps the error mode off the Windows Error Reporting dialog' {
            $header | Should Match '_set_abort_behavior\(0u,\s*_WRITE_ABORT_MSG \| _CALL_REPORTFAULT\)'
            $header | Should Match 'SetErrorMode\(GetErrorMode\(\) \| SEM_FAILCRITICALERRORS \| SEM_NOGPFAULTERRORBOX\)'
        }

        It 'installs a Debug-only report hook that leaves warnings to the CRT and ends every other report on stderr with abort' {
            $header | Should Match '(?s)#if defined\(_DEBUG\)\s+inline int __cdecl EndOnFailureReport\(int reportType, wchar_t\* message, int\* returnValue\) noexcept'
            $header | Should Match '(?s)reportType == _CRT_WARN\)\s*\{\s*return FALSE;'
            $header | Should Match '(?s)std::fputws\(.*?stderr\);\s*std::fflush\(stderr\);.*?std::abort\(\);'
            $header | Should Match '(?s)#if defined\(_DEBUG\)\s+static_cast<void>\(_CrtSetReportHookW2\(_CRT_RPTHOOK_INSTALL, &EndOnFailureReport\)\);\s+#endif'
        }
    }

    Context 'every native test entry routes first' {
        It 'finds the nine locally owned test executables and the PerformanceTests2 module' {
            @($nativeTestProjects | Where-Object { $_.Kind -eq 'Application' }).Count | Should Be 9
            @($nativeTestProjects | Where-Object { $_.Kind -eq 'DynamicLibrary' } | ForEach-Object { $_.Name }) | Should Be @('PerformanceTests2')
        }

        foreach ($project in $nativeTestProjects) {
            if ($project.Kind -eq 'Application') {
                It "$($project.Name) calls the routing as the first statement of its only entry point" {
                    @($project.Entries).Count | Should Be 1
                    $project.Entries[0].FirstStatement | Should BeExactly $routingCall
                }
            }
            elseif ($project.Kind -eq 'DynamicLibrary') {
                It "$($project.Name) routes in its module initializer, because its test host has no entry point of ours" {
                    @($project.Initializers).Count | Should Be 1
                    $project.Initializers[0].FirstStatement | Should BeExactly $routingCall
                }
            }
            else {
                It "$($project.Name) has a configuration type this contract knows how to check" {
                    $project.Kind | Should Match '^(Application|DynamicLibrary)$'
                }
            }
        }

        It 'includes the routing header from every file that calls it' {
            foreach ($project in $nativeTestProjects) {
                foreach ($file in (@($project.Entries) + @($project.Initializers) | ForEach-Object { $_.File } | Sort-Object -Unique)) {
                    $text = Get-RSRoutingSource -RelativePath (Join-Path (Join-Path 'Tests' $project.Name) $file)
                    $text | Should Match '#include "TestSupport/TestSupport\.Tests\.FailureReports\.h"'
                }
            }
        }
    }

    Context 'the product executables route only on a self-test path of a test-enabled build' {
        It 'routes RedSalamander.exe only for the in-product self-test suites' {
            $source = Get-RSRoutingSource -RelativePath 'RedSalamander\RedSalamander.cpp'
            ([regex]::Matches($source, [regex]::Escape($routingCall))).Count | Should Be 1

            $include = Get-RSRoutingContext -Source $source -Text '#include "TestSupport/TestSupport.Tests.FailureReports.h"'
            $include | Should Not BeNullOrEmpty
            @($include.Conditions) | Should Be @('ifdef ENABLE_TESTS')

            $call = Get-RSRoutingContext -Source $source -Text $routingCall
            $call | Should Not BeNullOrEmpty
            @($call.Conditions) | Should Be @('ifdef ENABLE_TESTS')
            $call.Before | Should Match 'hasArg\(L"--selftest"\) \|\| hasArg\(L"--compare-selftest"\) \|\| hasArg\(L"--commands-selftest"\) \|\| hasArg\(L"--fileops-selftest"\)'
        }


    }

    Context 'the failure-report self-test' {
        It 'is a hidden switch of ProductUiTests.exe that fails a Debug-CRT assertion on purpose and does nothing else' {
            $source = Get-RSRoutingSource -RelativePath 'Tests\ProductUiTests\ProductUi.Tests.Runner.cpp'
            $source | Should Match '(?s)int wmain\(int argc, wchar_t\*\* argv\)\s*\{\s*RedSalamander::TestSupport::RouteFailureReportsAwayFromDialogs\(\);'
            $source | Should Match 'std::wstring_view\(argv\[1\]\) == L"--failure-report-self-test"'
            $source | Should Match '(?s)#if defined\(_DEBUG\)\s+_ASSERTE\(! L"the failure-report self-test fails this check on purpose"\);\s+std::fputs\("The failed check returned instead of ending the run\.\\n", stderr\);\s+return 1;\s+#else'
            $source | Should Match '(?s)#else\s+std::fputs\("A Release build has no runtime checks to report\.\\n", stdout\);\s+return 0;\s+#endif'
        }

        # Debug and ASan Debug end as the CRT dialog's Abort would (exit code 3, the report on stderr); Release has no such checks.
        # A run that waits on a dialog is ended after the bound, and only that exact process is ended.
        $platforms = @('x64')
        if ($env:PROCESSOR_ARCHITECTURE -eq 'ARM64') {
            $platforms += 'ARM64'
        }
        $runs = @(foreach ($platform in $platforms) {
                foreach ($configuration in @('Debug', 'ASan Debug', 'Release')) {
                    $executable = Join-Path $repoRoot ('.build\{0}\{1}\ProductUiTests.exe' -f $platform, $configuration)
                    if ((Test-Path -LiteralPath $executable -PathType Leaf) -and (Test-RSExecutableContainsSelfTestSwitch -Path $executable)) {
                        [pscustomobject]@{ Platform = $platform; Configuration = $configuration; Path = $executable }
                    }
                }
            })

        if ($runs.Count -eq 0) {
            It 'ends a built ProductUiTests.exe run without waiting on a dialog (skipped: no built ProductUiTests.exe carries the switch yet)' -Skip {
            }
        }

        foreach ($run in $runs) {
            It "ends the $($run.Platform) $($run.Configuration) run within $($selfTestTimeoutMilliseconds / 1000) seconds without waiting on a dialog" {
                $result = Invoke-RSBoundedFailureReportRun -ExecutablePath $run.Path
                $result.Exited | Should Be $true
                if ($run.Configuration -eq 'Release') {
                    $result.ExitCode | Should Be 0
                }
                else {
                    $result.ExitCode | Should Be 3
                    $result.Output | Should Match 'failure-report self-test fails this check on purpose'
                    $result.Output | Should Not Match 'returned instead of ending'
                }
            }
        }
    }
}
