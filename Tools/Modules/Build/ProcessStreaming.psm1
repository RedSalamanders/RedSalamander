# Private build module; import through its reviewed callers.
Set-StrictMode -Version Latest

$sanitizedEnvironmentScript = Join-Path $PSScriptRoot 'SanitizedEnvironment.psm1'
if (-not (Test-Path $sanitizedEnvironmentScript)) {
    throw "Sanitized environment helper not found: $sanitizedEnvironmentScript"
}

Import-Module $sanitizedEnvironmentScript -Scope Local -ErrorAction Stop

function ConvertTo-RSStreamingQuotedArgument {
    param(
        [AllowNull()]
        [string]$Argument
    )

    if ($null -eq $Argument -or $Argument.Length -eq 0) {
        return '""'
    }

    if ($Argument -notmatch '[\s"]') {
        return $Argument
    }

    $builder = [System.Text.StringBuilder]::new()
    [void]$builder.Append('"')
    $backslashCount = 0
    foreach ($character in $Argument.ToCharArray()) {
        if ($character -eq '\') {
            $backslashCount++
            continue
        }

        if ($character -eq '"') {
            if ($backslashCount -gt 0) {
                [void]$builder.Append(('\' * ($backslashCount * 2)))
                $backslashCount = 0
            }

            [void]$builder.Append('\"')
            continue
        }

        if ($backslashCount -gt 0) {
            [void]$builder.Append(('\' * $backslashCount))
            $backslashCount = 0
        }

        [void]$builder.Append($character)
    }

    if ($backslashCount -gt 0) {
        [void]$builder.Append(('\' * ($backslashCount * 2)))
    }

    [void]$builder.Append('"')
    return $builder.ToString()
}

function Invoke-RSStreamingProcess {
    param(
        [Parameter(Mandatory = $true)]
        [string]$FilePath,

        [string[]]$Arguments = @(),

        [string]$WorkingDirectory = (Get-Location).Path,

        [string]$LogPath = '',

        [string]$ErrorLogPath = '',

        [hashtable]$AdditionalEnvironment = @{},

        [scriptblock]$OutputLineCallback,

        [ValidateRange(0, [int]::MaxValue)]
        [int]$TimeoutMilliseconds = 0,

        [ValidateRange(1, [int]::MaxValue)]
        [int]$OutputDrainTimeoutMilliseconds = 5000,

        [switch]$DelegateArtifactOperation
    )

    $psi = New-RSProcessStartInfo `
        -FilePath $FilePath `
        -Arguments $Arguments `
        -WorkingDirectory $WorkingDirectory `
        -AdditionalEnvironment $AdditionalEnvironment
    $psi.RedirectStandardOutput = $true
    $psi.RedirectStandardError = $true
    $psi.CreateNoWindow = $true

    # Resolve and compare both paths before creating either file. A same-path
    # configuration must fail without truncating the shared target.
    $resolvedLogPath = if ([string]::IsNullOrWhiteSpace($LogPath)) {
        ''
    }
    else {
        [System.IO.Path]::GetFullPath($LogPath)
    }
    $resolvedErrorLogPath = if ([string]::IsNullOrWhiteSpace($ErrorLogPath)) {
        ''
    }
    else {
        [System.IO.Path]::GetFullPath($ErrorLogPath)
    }
    if (-not [string]::IsNullOrWhiteSpace($resolvedLogPath) -and
        -not [string]::IsNullOrWhiteSpace($resolvedErrorLogPath) -and
        [string]::Equals($resolvedLogPath, $resolvedErrorLogPath, [System.StringComparison]::OrdinalIgnoreCase)) {
        throw 'LogPath and ErrorLogPath must name different files when stderr is split from stdout.'
    }

    $process = $null
    $stopwatch = [System.Diagnostics.Stopwatch]::StartNew()
    $logWriter = $null
    $errorLogWriter = $null

    try {
        if (-not [string]::IsNullOrWhiteSpace($resolvedLogPath)) {
            $logDir = Split-Path -Parent $resolvedLogPath
            if (-not [string]::IsNullOrWhiteSpace($logDir)) {
                [void](New-Item -ItemType Directory -Path $logDir -Force)
            }

            $encoding = [System.Text.UTF8Encoding]::new($false)
            $logWriter = [System.IO.StreamWriter]::new($resolvedLogPath, $false, $encoding)
            $logWriter.AutoFlush = $true
        }
        if (-not [string]::IsNullOrWhiteSpace($resolvedErrorLogPath)) {
            $errorLogDir = Split-Path -Parent $resolvedErrorLogPath
            if (-not [string]::IsNullOrWhiteSpace($errorLogDir)) {
                [void](New-Item -ItemType Directory -Path $errorLogDir -Force)
            }

            $encoding = [System.Text.UTF8Encoding]::new($false)
            $errorLogWriter = [System.IO.StreamWriter]::new($resolvedErrorLogPath, $false, $encoding)
            $errorLogWriter.AutoFlush = $true
        }

        $process = Start-RSContainedProcess `
            -ProcessStartInfo $psi `
            -DelegateArtifactOperation:$DelegateArtifactOperation

        $stdOutOpen = $true
        $stdErrOpen = $true
        $stdOutTask = $process.StandardOutput.ReadLineAsync()
        $stdErrTask = $process.StandardError.ReadLineAsync()
        $exitCode = $null
        $exitObservedAtMilliseconds = 0L

        while ($true) {
            if ($null -eq $exitCode -and $process.HasExited) {
                $exitCode = [int]$process.ExitCode
                $exitObservedAtMilliseconds = $stopwatch.ElapsedMilliseconds
            }

            if ($null -eq $exitCode -and
                $TimeoutMilliseconds -gt 0 -and
                $stopwatch.ElapsedMilliseconds -ge $TimeoutMilliseconds) {
                $processId = $process.Id
                Close-RSContainedProcess -Process $process
                $process = $null
                throw "Timed out after $TimeoutMilliseconds milliseconds running '$FilePath' (PID $processId)."
            }

            if ($null -ne $exitCode -and
                ($stdOutOpen -or $stdErrOpen) -and
                ($stopwatch.ElapsedMilliseconds - $exitObservedAtMilliseconds) -ge $OutputDrainTimeoutMilliseconds) {
                $diagnostic = "[process-streaming] Parent process exited with code $exitCode, but redirected output remained open for $OutputDrainTimeoutMilliseconds milliseconds; closing the containment job to release inherited pipe handles."
                $diagnosticWriter = if ($null -ne $errorLogWriter) { $errorLogWriter } else { $logWriter }
                if ($null -ne $diagnosticWriter) {
                    $diagnosticWriter.WriteLine($diagnostic)
                }
                Write-Verbose $diagnostic

                Close-RSContainedProcess -Process $process
                $process = $null

                # Closing the job releases handles inherited by descendants. Give
                # each already-pending ReadLineAsync a final bounded chance to
                # return buffered text, including a last line without a newline.
                $finalDrain = [System.Diagnostics.Stopwatch]::StartNew()
                while (($stdOutOpen -or $stdErrOpen) -and
                    $finalDrain.ElapsedMilliseconds -lt $OutputDrainTimeoutMilliseconds) {
                    $pendingTasks = [System.Collections.Generic.List[System.Threading.Tasks.Task[string]]]::new()
                    if ($stdOutOpen) {
                        [void]$pendingTasks.Add($stdOutTask)
                    }
                    if ($stdErrOpen) {
                        [void]$pendingTasks.Add($stdErrTask)
                    }
                    if ($pendingTasks.Count -eq 0) {
                        break
                    }

                    $remainingDrainMilliseconds = [Math]::Max(
                        1,
                        $OutputDrainTimeoutMilliseconds - [int]$finalDrain.ElapsedMilliseconds)
                    $completedDrainIndex = [System.Threading.Tasks.Task]::WaitAny(
                        $pendingTasks.ToArray(),
                        [Math]::Min(50, $remainingDrainMilliseconds))
                    if ($completedDrainIndex -lt 0) {
                        continue
                    }

                    $completedDrainTask = $pendingTasks[$completedDrainIndex]
                    $drainedIsError = $stdErrOpen -and
                        [object]::ReferenceEquals($completedDrainTask, $stdErrTask)
                    try {
                        $drainedLine = $completedDrainTask.GetAwaiter().GetResult()
                    }
                    catch {
                        $drainFailure = "[process-streaming] Final redirected-output drain ended with '$($_.Exception.Message)'."
                        $drainFailureWriter = if ($null -ne $errorLogWriter) { $errorLogWriter } else { $logWriter }
                        if ($null -ne $drainFailureWriter) {
                            $drainFailureWriter.WriteLine($drainFailure)
                        }
                        Write-Verbose $drainFailure
                        $drainedLine = $null
                    }

                    if ($drainedIsError) {
                        $stdErrOpen = $false
                    }
                    else {
                        $stdOutOpen = $false
                    }
                    if ($null -eq $drainedLine) {
                        continue
                    }

                    $drainedWriter = if ($drainedIsError -and $null -ne $errorLogWriter) {
                        $errorLogWriter
                    }
                    else {
                        $logWriter
                    }
                    if ($null -ne $drainedWriter) {
                        $drainedWriter.WriteLine($drainedLine)
                    }
                    if ($OutputLineCallback) {
                        & $OutputLineCallback $drainedLine $drainedIsError
                    }
                    else {
                        Write-Host $drainedLine
                    }
                }
                $finalDrain.Stop()
                if ($stdOutOpen -or $stdErrOpen) {
                    $finalDrainTimeout = "[process-streaming] Final redirected-output drain reached its $OutputDrainTimeoutMilliseconds millisecond bound."
                    $finalDrainTimeoutWriter = if ($null -ne $errorLogWriter) { $errorLogWriter } else { $logWriter }
                    if ($null -ne $finalDrainTimeoutWriter) {
                        $finalDrainTimeoutWriter.WriteLine($finalDrainTimeout)
                    }
                    Write-Verbose $finalDrainTimeout
                }
                break
            }

            $tasks = [System.Collections.Generic.List[System.Threading.Tasks.Task[string]]]::new()
            if ($stdOutOpen) {
                [void]$tasks.Add($stdOutTask)
            }
            if ($stdErrOpen) {
                [void]$tasks.Add($stdErrTask)
            }

            if ($tasks.Count -eq 0) {
                if ($null -ne $exitCode) {
                    break
                }

                Start-Sleep -Milliseconds 25
                continue
            }

            $completedIndex = [System.Threading.Tasks.Task]::WaitAny($tasks.ToArray(), 50)
            if ($completedIndex -lt 0) {
                if ($process.HasExited -and -not $stdOutOpen -and -not $stdErrOpen) {
                    break
                }

                continue
            }

            $completedTask = $tasks[$completedIndex]
            $isError = $stdErrOpen -and [object]::ReferenceEquals($completedTask, $stdErrTask)
            $line = $completedTask.GetAwaiter().GetResult()

            if ($null -eq $line) {
                if ($isError) {
                    $stdErrOpen = $false
                }
                else {
                    $stdOutOpen = $false
                }
            }
            else {
                $lineWriter = if ($isError -and $null -ne $errorLogWriter) {
                    $errorLogWriter
                }
                else {
                    $logWriter
                }
                if ($null -ne $lineWriter) {
                    $lineWriter.WriteLine($line)
                }

                if ($OutputLineCallback) {
                    & $OutputLineCallback $line $isError
                }
                else {
                    Write-Host $line
                }

                if ($isError) {
                    $stdErrTask = $process.StandardError.ReadLineAsync()
                }
                else {
                    $stdOutTask = $process.StandardOutput.ReadLineAsync()
                }
            }

            if ($null -ne $exitCode -and -not $stdOutOpen -and -not $stdErrOpen) {
                break
            }
        }

        if ($null -eq $exitCode) {
            throw "Process streaming ended without observing an exit code for '$FilePath'."
        }

        $global:LASTEXITCODE = $exitCode
        return $exitCode
    }
    finally {
        $stopwatch.Stop()
        Close-RSContainedProcess -Process $process

        if ($logWriter) {
            $logWriter.Dispose()
        }
        if ($errorLogWriter) {
            $errorLogWriter.Dispose()
        }
    }
}

Export-ModuleMember -Function @(
    'ConvertTo-RSStreamingQuotedArgument',
    'Invoke-RSStreamingProcess'
)
