param(
    [string]$Root = (Resolve-Path "$PSScriptRoot\..").Path
)

$allowedPathPatterns = @(
    '\\SelfTest\\',
    '\\Tests\\',
    '\\Scripts\\VerifyNoProductionGetCursorPos\.ps1$'
)

$requiredDiagnosticAnnotation = '// getcursorpos-allow: diagnostic-only'

$files = @(
    (Join-Path $Root 'Common'),
    (Join-Path $Root 'RedSalamander'),
    (Join-Path $Root 'Plugins')
)

$violations = New-Object System.Collections.Generic.List[string]
foreach ($fileRoot in $files) {
    if (-not (Test-Path -LiteralPath $fileRoot)) {
        continue
    }

    foreach ($item in Get-ChildItem -Path $fileRoot -Recurse -Include *.cpp,*.h) {
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

if ($violations.Count -gt 0) {
    Write-Host 'Production GetCursorPos violations:'
    $violations | ForEach-Object { Write-Host $_ }
    exit 1
}

Write-Host 'No production GetCursorPos violations found.'
exit 0
