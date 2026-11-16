# Internal engine-free Git line/stdin transport. Import through its two testing callers.
Set-StrictMode -Version Latest

function Invoke-RSTestGitMetadataCapture {
    param([string]$Root, [string[]]$Arguments, [AllowEmptyCollection()][string[]]$InputLines)
    # Engine-free bootstrap projection: native pipeline text/stdin semantics differ
    # from the engine-dependent streaming factory. No public process surface.
    $start = [Diagnostics.ProcessStartInfo]::new('git')
    $start.UseShellExecute = $false
    $start.CreateNoWindow = $true
    $start.RedirectStandardOutput = $true
    $start.RedirectStandardError = $true
    $start.StandardOutputEncoding = [Text.UTF8Encoding]::new($false)
    $start.StandardErrorEncoding = [Console]::OutputEncoding
    if ($PSBoundParameters.ContainsKey('InputLines')) {
        $start.RedirectStandardInput = $true
        $start.StandardInputEncoding = $OutputEncoding
    }
    foreach ($argument in @('-C', $Root) + $Arguments) { [void]$start.ArgumentList.Add($argument) }
    $process = [Diagnostics.Process]::new()
    $process.StartInfo = $start
    try {
        if (-not $process.Start()) { throw 'Unable to start Git metadata query.' }
        $stdoutTask = $process.StandardOutput.ReadToEndAsync()
        $stderrTask = $process.StandardError.ReadToEndAsync()
        $inputError = $null
        if ($start.RedirectStandardInput) {
            try { foreach ($line in $InputLines) { $process.StandardInput.WriteLine($line) } }
            catch [IO.IOException] { $inputError = $_ }
            finally { try { $process.StandardInput.Close() } catch [IO.IOException] { if (-not $inputError) { $inputError = $_ } } }
        }
        $process.WaitForExit()
        $stdout = $stdoutTask.GetAwaiter().GetResult()
        $stderr = $stderrTask.GetAwaiter().GetResult()
        $global:LASTEXITCODE = $process.ExitCode
        if ($inputError -and $process.ExitCode -eq 0) { throw $inputError }
        $lines = [Collections.Generic.List[string]]::new()
        $reader = [IO.StringReader]::new($stdout)
        try { while ($null -ne ($line = $reader.ReadLine())) { $lines.Add($line) } }
        finally { $reader.Dispose() }
        return [pscustomobject]@{ ExitCode = $process.ExitCode; Output = [string[]]$lines.ToArray(); ErrorText = $stderr }
    } finally { $process.Dispose() }
}

Export-ModuleMember -Function Invoke-RSTestGitMetadataCapture
