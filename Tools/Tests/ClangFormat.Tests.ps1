Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$rsFormatterRepository = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
Import-Module (Join-Path $rsFormatterRepository 'Tools/Modules/Testing/TestSandbox.psm1') -Force
Import-Module (Join-Path $rsFormatterRepository 'Tools/Modules/Build/ClangFormat.psm1') -Force

InModuleScope ClangFormat {
    function New-FormatterFixture {
        param([string[]]$Entries = @('clang_format/data/bin/clang-format.exe', '../unrelated.txt'))
        $repository = [IO.Path]::GetFullPath((Join-Path (Split-Path (Get-Module ClangFormat).Path) '../../..'))
        $root = New-RSTestSandboxScratchDirectory -RepoRoot $repository -Harness tools-pester -Case ('formatter-' + [guid]::NewGuid().ToString('N'))
        [void][IO.Directory]::CreateDirectory((Join-Path $root 'Build'))
        $wheel = Join-Path $root 'fixture.whl'
        $archive = [IO.Compression.ZipFile]::Open($wheel, [IO.Compression.ZipArchiveMode]::Create)
        try {
            foreach ($name in $Entries) {
                $writer = [IO.StreamWriter]::new($archive.CreateEntry($name).Open())
                try { $writer.Write('verified native fixture') } finally { $writer.Dispose() }
            }
        } finally { $archive.Dispose() }
        $pin = @{version='22.1.3'; wheelSha256=(Get-FileHash -LiteralPath $wheel -Algorithm SHA256).Hash.ToLowerInvariant()}
        [IO.File]::WriteAllText((Join-Path $root 'Build/clang-format.json'), ($pin | ConvertTo-Json))
        return [pscustomobject]@{Root=$root; Wheel=$wheel; Output=(Join-Path $root '.build/format/clang_format/data/bin/clang-format.exe')}
    }

    Describe 'Pinned native formatter acquisition' {
        BeforeEach {
            $script:formatterFixture = New-FormatterFixture
            Mock Assert-RSClangFormatVersion {}
            Mock Invoke-RestMethod { throw 'Unexpected network access' }
            Mock Invoke-WebRequest { throw 'Unexpected network access' }
        }
        It 'supports WhatIf without downloading or creating the install directory' {
            Install-RSClangFormat -RepositoryRoot $script:formatterFixture.Root -WhatIf
            (Test-Path (Join-Path $script:formatterFixture.Root '.build')) | Should Be $false
            Assert-MockCalled Invoke-RestMethod -Times 0 -Exactly
        }
        It 'extracts only the native executable from a verified offline wheel' {
            $result = Install-RSClangFormat -RepositoryRoot $script:formatterFixture.Root -WheelPath $script:formatterFixture.Wheel
            $result | Should Be $script:formatterFixture.Output
            [IO.File]::ReadAllText($result) | Should Be 'verified native fixture'
            @(Get-ChildItem -LiteralPath (Join-Path $script:formatterFixture.Root '.build') -File -Recurse).Count | Should Be 1
            Assert-MockCalled Assert-RSClangFormatVersion -Times 1 -Exactly
        }
        It 'rejects a hash mismatch and preserves an existing formatter and the caller wheel' {
            [void][IO.Directory]::CreateDirectory((Split-Path $script:formatterFixture.Output))
            [IO.File]::WriteAllText($script:formatterFixture.Output, 'previous formatter')
            [IO.File]::AppendAllText($script:formatterFixture.Wheel, 'tampered')
            { Install-RSClangFormat -RepositoryRoot $script:formatterFixture.Root -WheelPath $script:formatterFixture.Wheel } | Should Throw 'pinned SHA-256'
            [IO.File]::ReadAllText($script:formatterFixture.Output) | Should Be 'previous formatter'
            (Test-Path $script:formatterFixture.Wheel) | Should Be $true
        }
        It 'rejects a missing native entry' {
            $fixture = New-FormatterFixture -Entries @('other.txt')
            { Install-RSClangFormat -RepositoryRoot $fixture.Root -WheelPath $fixture.Wheel } | Should Throw 'exactly one native executable'
            (Test-Path $fixture.Output) | Should Be $false
        }
        It 'rejects duplicate native entries' {
            $fixture = New-FormatterFixture -Entries @('clang_format/data/bin/clang-format.exe', 'clang_format/data/bin/clang-format.exe')
            { Install-RSClangFormat -RepositoryRoot $fixture.Root -WheelPath $fixture.Wheel } | Should Throw 'exactly one native executable'
        }
        It 'preserves the old executable when version validation fails' {
            [void][IO.Directory]::CreateDirectory((Split-Path $script:formatterFixture.Output))
            [IO.File]::WriteAllText($script:formatterFixture.Output, 'previous formatter')
            Mock Assert-RSClangFormatVersion { throw 'wrong native version' }
            { Install-RSClangFormat -RepositoryRoot $script:formatterFixture.Root -WheelPath $script:formatterFixture.Wheel } | Should Throw 'wrong native version'
            [IO.File]::ReadAllText($script:formatterFixture.Output) | Should Be 'previous formatter'
            @(Get-ChildItem -LiteralPath (Split-Path $script:formatterFixture.Output) -File).Count | Should Be 1
        }
        It 'checks cached executable bytes and repairs a modified executable' {
            $null = Install-RSClangFormat -RepositoryRoot $script:formatterFixture.Root -WheelPath $script:formatterFixture.Wheel
            [IO.File]::WriteAllText($script:formatterFixture.Output, 'modified cache')
            $null = Install-RSClangFormat -RepositoryRoot $script:formatterFixture.Root -WheelPath $script:formatterFixture.Wheel
            [IO.File]::ReadAllText($script:formatterFixture.Output) | Should Be 'verified native fixture'
        }
        It 'downloads once, verifies the package, and subsequently reuses the verified cache' {
            Mock Invoke-RestMethod {
                $hash = (Get-FileHash -LiteralPath $script:formatterFixture.Wheel -Algorithm SHA256).Hash.ToLowerInvariant()
                return @{files=@(@{url='https://packages.example.invalid/fixture.whl'; hashes=@{sha256=$hash}})}
            }
            Mock Invoke-WebRequest { param($Uri,$OutFile) [IO.File]::Copy($script:formatterFixture.Wheel,$OutFile) }
            $null = Install-RSClangFormat -RepositoryRoot $script:formatterFixture.Root
            $null = Install-RSClangFormat -RepositoryRoot $script:formatterFixture.Root
            Assert-MockCalled Invoke-RestMethod -Times 1 -Exactly
            Assert-MockCalled Invoke-WebRequest -Times 1 -Exactly
            [IO.File]::ReadAllText($script:formatterFixture.Output) | Should Be 'verified native fixture'
        }
        It 'rejects a corrupt download before publishing a cached wheel or executable' {
            Mock Invoke-RestMethod {
                $hash = (Get-FileHash -LiteralPath $script:formatterFixture.Wheel -Algorithm SHA256).Hash.ToLowerInvariant()
                return @{files=@(@{url='https://packages.example.invalid/fixture.whl'; hashes=@{sha256=$hash}})}
            }
            Mock Invoke-WebRequest { param($Uri,$OutFile) [IO.File]::WriteAllText($OutFile,'corrupt') }
            { Install-RSClangFormat -RepositoryRoot $script:formatterFixture.Root } | Should Throw 'pinned SHA-256'
            (Test-Path $script:formatterFixture.Output) | Should Be $false
            @(Get-ChildItem -LiteralPath (Join-Path $script:formatterFixture.Root '.build/format') -File).Count | Should Be 0
        }
    }
}
