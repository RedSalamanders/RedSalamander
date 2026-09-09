$repoRoot = (Resolve-Path (Join-Path $PSScriptRoot '..\..')).Path
$modulePath = Join-Path $repoRoot 'Tools\TerminalEngine\TerminalEnginePeIdentity.psm1'
Import-Module $modulePath -Force

function Get-TestPePath {
    $pwshPath = (Get-Process -Id $PID).Path
    if (-not [string]::IsNullOrWhiteSpace($pwshPath) -and
        [IO.File]::Exists($pwshPath)) {
        return $pwshPath
    }
    return (Join-Path $env:SystemRoot 'System32\notepad.exe')
}

Describe 'Terminal engine PE semantic identity' {
    It 'parses a native PE into closed, sorted identity records' {
        $identity = Get-RSPeIdentity -Path (Get-TestPePath)

        $identity.sha256 | Should Match '^[0-9a-f]{64}$'
        $identity.normalizedSha256 | Should Match '^[0-9a-f]{64}$'
        (@('AMD64', 'ARM64') -contains $identity.machine) | Should Be $true
        $identity.semantic.formatId | Should Be 'red-salamander-pe-semantic-identity'
        $identity.semantic.version | Should Be 1
        $identity.semantic.sections.Count | Should BeGreaterThan 0

        $importKeys = @($identity.semantic.imports | ForEach-Object {
            '{0}{1}{2}' -f $_.module, [char]0, $_.symbol
        })
        $sortedImportKeys = @($importKeys | Sort-Object -CaseSensitive)
        ($importKeys -join "`n") | Should Be ($sortedImportKeys -join "`n")

        $exportKeys = @($identity.semantic.exports | ForEach-Object {
            '{0}{1}{2:D10}' -f $(if ($null -eq $_.name) { '' } else { $_.name }),
                [char]0,
                [uint64]$_.ordinal
        })
        $sortedExportKeys = @($exportKeys | Sort-Object -CaseSensitive)
        ($exportKeys -join "`n") | Should Be ($sortedExportKeys -join "`n")
    }

    It 'rejects malformed and truncated inputs' {
        $malformed = Join-Path $TestDrive 'not-pe.bin'
        [IO.File]::WriteAllBytes($malformed, [byte[]](0x4d, 0x5a, 0, 0))

        $message = $null
        try {
            Get-RSPeIdentity -Path $malformed | Out-Null
        } catch {
            $message = $_.Exception.Message
        }
        $message | Should Match 'not a PE image'
    }

    It 'normalizes the COFF timestamp while retaining the exact packaged hash' {
        $source = Get-TestPePath
        $copy = Join-Path $TestDrive 'timestamp-mutated.exe'
        [IO.File]::Copy($source, $copy)

        $before = Get-RSPeIdentity -Path $copy
        $beforeSemantic = Get-RSPeSemanticSha256 -Path $copy
        $bytes = [IO.File]::ReadAllBytes($copy)
        $peOffset = [BitConverter]::ToInt32($bytes, 0x3c)
        $timestampOffset = $peOffset + 8
        $originalTimestamp = [BitConverter]::ToUInt32($bytes, $timestampOffset)
        $replacementTimestamp = $originalTimestamp -bxor [uint32]0x5a5aa5a5
        [BitConverter]::GetBytes($replacementTimestamp).CopyTo($bytes, $timestampOffset)
        [IO.File]::WriteAllBytes($copy, $bytes)

        $after = Get-RSPeIdentity -Path $copy
        $afterSemantic = Get-RSPeSemanticSha256 -Path $copy
        $after.sha256 | Should Not Be $before.sha256
        $after.normalizedSha256 | Should Be $before.normalizedSha256
        $afterSemantic | Should Be $beforeSemantic
        $after.semantic | ConvertTo-Json -Depth 20 -Compress |
            Should Be ($before.semantic | ConvertTo-Json -Depth 20 -Compress)
    }

    It 'produces a lowercase canonical semantic digest' {
        Get-RSPeSemanticSha256 -Path (Get-TestPePath) |
            Should Match '^[0-9a-f]{64}$'
    }
}
