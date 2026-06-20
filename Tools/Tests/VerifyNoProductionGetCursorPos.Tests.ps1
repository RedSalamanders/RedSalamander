Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)

$Root = $repoRoot

$allowedPathPatterns = @(
    '\\SelfTest\\',
    '\\Tests\\',
    '\\Tools\\Tests\\VerifyNoProductionGetCursorPos\.Tests\.ps1$'
)
$requiredDiagnosticAnnotation = 'getcursorpos-allow:'
$files = @(
    (Join-Path $Root 'Common'),
    (Join-Path $Root 'RedSalamander'),
    (Join-Path $Root 'Plugins')
)

function Get-RSGetCursorPosViolations {
    $violations = New-Object System.Collections.Generic.List[string]
    foreach ($fileRoot in $files) {
        if (-not (Test-Path -LiteralPath $fileRoot)) {
            continue
        }

        foreach ($item in Get-ChildItem -Path $fileRoot -Recurse -File -Include *.cpp,*.h) {
            $path = $item.FullName
            $relative = $path.Substring($Root.Length).TrimStart('\')
            if ($allowedPathPatterns | Where-Object { $relative -match $_ }) {
                continue
            }

            $lineNo = 0
            foreach ($line in Get-Content -Path $path) {
                $lineNo++
                if ($line -notmatch 'GetCursorPos\s*\(') {
                    continue
                }

                if ($line.Contains($requiredDiagnosticAnnotation)) {
                    continue
                }

                $violations.Add("${relative}:${lineNo}: $line")
            }
        }
    }

    # The pipeline unrolls the list; the @() at the call site rebuilds a
    # (possibly empty) array so .Count stays valid under StrictMode.
    return $violations
}

Describe 'VerifyNoProductionGetCursorPos' {
    It 'finds no production GetCursorPos() usage outside SelfTest/Tests sources' {
        $violations = @(Get-RSGetCursorPosViolations)
        if ($violations.Count -gt 0) {
            Write-Host 'Production GetCursorPos violations:'
            $violations | ForEach-Object { Write-Host $_ }
        }
        $violations.Count | Should Be 0
    }
}
