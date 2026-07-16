Set-StrictMode -Version Latest

$sanitizedEnvironmentScript = Join-Path $PSScriptRoot 'SanitizedEnvironment.ps1'
if (-not (Test-Path $sanitizedEnvironmentScript)) {
    throw "Sanitized environment helper not found: $sanitizedEnvironmentScript"
}

. $sanitizedEnvironmentScript

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

        [hashtable]$AdditionalEnvironment = @{},

        [scriptblock]$OutputLineCallback
    )

    $psi = New-RSProcessStartInfo `
        -FilePath $FilePath `
        -Arguments $Arguments `
        -WorkingDirectory $WorkingDirectory `
        -AdditionalEnvironment $AdditionalEnvironment
    $psi.RedirectStandardOutput = $true
    $psi.RedirectStandardError = $true
    $psi.CreateNoWindow = $true

    $process = $null

    $logWriter = $null
    if (-not [string]::IsNullOrWhiteSpace($LogPath)) {
        $logDir = Split-Path -Parent $LogPath
        if (-not [string]::IsNullOrWhiteSpace($logDir)) {
            [void](New-Item -ItemType Directory -Path $logDir -Force)
        }

        $encoding = [System.Text.UTF8Encoding]::new($false)
        $logWriter = [System.IO.StreamWriter]::new($LogPath, $false, $encoding)
        $logWriter.AutoFlush = $true
    }

    try {
        $process = Start-RSContainedProcess -ProcessStartInfo $psi

        $stdOutOpen = $true
        $stdErrOpen = $true
        $stdOutTask = $process.StandardOutput.ReadLineAsync()
        $stdErrTask = $process.StandardError.ReadLineAsync()

        while ($true) {
            $tasks = [System.Collections.Generic.List[System.Threading.Tasks.Task[string]]]::new()
            if ($stdOutOpen) {
                [void]$tasks.Add($stdOutTask)
            }
            if ($stdErrOpen) {
                [void]$tasks.Add($stdErrTask)
            }

            if ($tasks.Count -eq 0) {
                if ($process.HasExited) {
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
                if ($logWriter) {
                    $logWriter.WriteLine($line)
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

            if ($process.HasExited -and -not $stdOutOpen -and -not $stdErrOpen) {
                break
            }
        }

        $process.WaitForExit()
        $global:LASTEXITCODE = $process.ExitCode
        return $process.ExitCode
    }
    finally {
        Close-RSContainedProcess -Process $process

        if ($logWriter) {
            $logWriter.Dispose()
        }

    }
}
