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

    function Remove-FormatterFixtureJunction {
        param([string]$Path, [string]$ExpectedTarget)
        # The canonical tree cleaner rejects absolute links and removes a whole tree.
        # This fixture needs to unlink exactly its owned junction while retaining all witnesses.
        $entry = Get-Item -LiteralPath $Path -Force -ErrorAction SilentlyContinue
        if ($null -eq $entry -or ($entry.Attributes -band [IO.FileAttributes]::ReparsePoint) -eq 0) { return }
        if ($entry.LinkType -cne 'Junction' -or @($entry.Target).Count -ne 1 -or
            -not [string]::Equals([string]$entry.Target, $ExpectedTarget, [StringComparison]::OrdinalIgnoreCase)) {
            throw "Refusing to unlink a junction the formatter fixture does not own: $Path"
        }
        [IO.Directory]::Delete($Path, $false)
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
            Assert-MockCalled Invoke-RestMethod -Times 0 -Exactly -Scope It
        }
        It 'extracts only the native executable from a verified offline wheel' {
            $result = Install-RSClangFormat -RepositoryRoot $script:formatterFixture.Root -WheelPath $script:formatterFixture.Wheel
            $result | Should Be $script:formatterFixture.Output
            [IO.File]::ReadAllText($result) | Should Be 'verified native fixture'
            @(Get-ChildItem -LiteralPath (Join-Path $script:formatterFixture.Root '.build') -File -Recurse -Force | Where-Object Name -NE '.install.lock').Count | Should Be 1
            Assert-MockCalled Assert-RSClangFormatVersion -Times 1 -Exactly -Scope It
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
            Assert-MockCalled Invoke-RestMethod -Times 1 -Exactly -Scope It
            Assert-MockCalled Invoke-WebRequest -Times 1 -Exactly -Scope It
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
            @(Get-ChildItem -LiteralPath (Join-Path $script:formatterFixture.Root '.build/format') -File -Force | Where-Object Name -NE '.install.lock').Count | Should Be 0
        }
        It 'rejects a junction at <RelativePath> before writing or executing' -TestCases @(
            @{RelativePath='.build'},
            @{RelativePath='.build/format'},
            @{RelativePath='.build/format/clang_format/data/bin'}
        ) {
            param($RelativePath)
            $target = $script:formatterFixture.Root + '-redirected'
            [void][IO.Directory]::CreateDirectory($target)
            $sentinel = Join-Path $target 'preserve.txt'
            [IO.File]::WriteAllText($sentinel, 'untouched target')
            $before = (Get-FileHash -LiteralPath $sentinel).Hash
            $link = Join-Path $script:formatterFixture.Root $RelativePath
            [void][IO.Directory]::CreateDirectory((Split-Path $link))
            try {
                $null = New-RSTestSandboxLink -Path $link -Target $target -ItemType Junction
                { Install-RSClangFormat -RepositoryRoot $script:formatterFixture.Root -WheelPath $script:formatterFixture.Wheel } | Should Throw 'reparse point'
                Assert-MockCalled Assert-RSClangFormatVersion -Times 0 -Exactly -Scope It
                Assert-MockCalled Invoke-RestMethod -Times 0 -Exactly -Scope It
            } finally { Remove-FormatterFixtureJunction -Path $link -ExpectedTarget $target }
            (Test-Path -LiteralPath $link) | Should Be $false
            (Get-FileHash -LiteralPath $sentinel).Hash | Should Be $before
            @(Get-ChildItem -LiteralPath $target -Force).Count | Should Be 1
        }
        It 'revalidates the output chain at publication and refuses redirected cleanup' {
            [void][IO.Directory]::CreateDirectory((Split-Path $script:formatterFixture.Output))
            [IO.File]::WriteAllText($script:formatterFixture.Output, 'previous formatter')
            $script:formatterRedirectedTarget = $script:formatterFixture.Root + '-publication-target'
            [void][IO.Directory]::CreateDirectory($script:formatterRedirectedTarget)
            Mock Assert-RSClangFormatVersion {
                param($Path, $Version)
                $parent = Split-Path $Path
                [IO.Directory]::Move($parent, ($parent + '.original'))
                $null = New-RSTestSandboxLink -Path $parent -Target $script:formatterRedirectedTarget -ItemType Junction
                [IO.File]::WriteAllText((Join-Path $script:formatterRedirectedTarget ([IO.Path]::GetFileName($Path))), 'preserve staged-name target')
                [IO.File]::WriteAllText((Join-Path $script:formatterRedirectedTarget 'clang-format.exe'), 'preserve published-name target')
            }
            $link = Split-Path $script:formatterFixture.Output
            try {
                { Install-RSClangFormat -RepositoryRoot $script:formatterFixture.Root -WheelPath $script:formatterFixture.Wheel -WarningAction SilentlyContinue } | Should Throw 'reparse point'
            } finally { Remove-FormatterFixtureJunction -Path $link -ExpectedTarget $script:formatterRedirectedTarget }
            (Test-Path -LiteralPath $link) | Should Be $false
            $original = Join-Path ((Split-Path $script:formatterFixture.Output) + '.original') 'clang-format.exe'
            [IO.File]::ReadAllText($original) | Should Be 'previous formatter'
            $targetFiles = @(Get-ChildItem -LiteralPath $script:formatterRedirectedTarget -File -Force)
            $targetFiles.Count | Should Be 2
            [IO.File]::ReadAllText((Join-Path $script:formatterRedirectedTarget 'clang-format.exe')) | Should Be 'preserve published-name target'
            @($targetFiles | Where-Object { [IO.File]::ReadAllText($_.FullName) -eq 'preserve staged-name target' }).Count | Should Be 1
        }
        It 'resolves an explicit relative wheel from the PowerShell location once' {
            $other = $script:formatterFixture.Root + '-process-cwd'
            [void][IO.Directory]::CreateDirectory($other)
            $otherWheel = Join-Path $other 'fixture.whl'
            [IO.File]::WriteAllText($otherWheel, 'different process-relative input')
            $previousProcessDirectory = [Environment]::CurrentDirectory
            Push-Location -LiteralPath $script:formatterFixture.Root
            try {
                [Environment]::CurrentDirectory = $other
                $result = Install-RSClangFormat -RepositoryRoot '.' -WheelPath './fixture.whl'
                $result | Should Be $script:formatterFixture.Output
                [IO.File]::ReadAllText($result) | Should Be 'verified native fixture'
                [IO.File]::ReadAllText($otherWheel) | Should Be 'different process-relative input'
            } finally {
                [Environment]::CurrentDirectory = $previousProcessDirectory
                Pop-Location
            }
        }
        It 'keeps the verified wheel open against modification until extraction and publication finish' {
            Mock Assert-RSClangFormatVersion {
                $null = { [IO.File]::WriteAllText($script:formatterFixture.Wheel, 'replacement') } | Should Throw 'used by another process'
            }
            $result = Install-RSClangFormat -RepositoryRoot $script:formatterFixture.Root -WheelPath $script:formatterFixture.Wheel
            [IO.File]::ReadAllText($result) | Should Be 'verified native fixture'
            Assert-MockCalled Assert-RSClangFormatVersion -Times 1 -Exactly -Scope It
        }
        It 'fails immediately on a competing cache lease before downloading or replacing files' {
            $cache = Join-Path $script:formatterFixture.Root '.build/format'
            [void][IO.Directory]::CreateDirectory($cache)
            $lease = [IO.File]::Open((Join-Path $cache '.install.lock'), [IO.FileMode]::OpenOrCreate, [IO.FileAccess]::ReadWrite, [IO.FileShare]::None)
            try {
                { Install-RSClangFormat -RepositoryRoot $script:formatterFixture.Root } | Should Throw 'used by another process'
                (Test-Path $script:formatterFixture.Output) | Should Be $false
                Assert-MockCalled Invoke-RestMethod -Times 0 -Exactly -Scope It
                Assert-MockCalled Assert-RSClangFormatVersion -Times 0 -Exactly -Scope It
            } finally { $lease.Dispose() }
        }
    }
}
