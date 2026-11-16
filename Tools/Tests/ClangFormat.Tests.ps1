Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$rsFormatterRepository = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
Import-Module (Join-Path $rsFormatterRepository 'Tools/Modules/Testing/TestSandbox.psm1') -Force
Import-Module (Join-Path $rsFormatterRepository 'Tools/Modules/Build/ClangFormat.psm1') -Force
Import-Module (Join-Path $PSScriptRoot 'TestSupport.psm1') -ErrorAction Stop

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

    Describe 'Engine-free native formatter capture' {
        It 'rejects unsafe capture flags in the actual transport before launching a child' {
            $body = (Get-Command Invoke-RSClangFormatCapture).ScriptBlock.ToString()
            $validate = {
                param($text)
                $start = $text.IndexOf('$process.Start()', [StringComparison]::Ordinal)
                if ($start -lt 0) { throw 'Missing formatter child start.' }
                foreach ($required in @('$start.UseShellExecute = $false', '$start.CreateNoWindow = $true',
                        '$start.RedirectStandardOutput = $true', '$start.RedirectStandardError = $true')) {
                    $position = $text.IndexOf($required, [StringComparison]::Ordinal)
                    if ($position -lt 0 -or $position -ge $start) { throw 'Unsafe formatter child configuration.' }
                }
            }
            & $validate $body
            foreach ($mutation in @(
                    @('$start.CreateNoWindow = $true', '$start.CreateNoWindow = $false'),
                    @('$start.UseShellExecute = $false', '$start.UseShellExecute = $true'),
                    @('$start.RedirectStandardError = $true', '$start.RedirectStandardError = $false'))) {
                { & $validate ($body.Replace($mutation[0], $mutation[1])) } | Should Throw 'Unsafe formatter child configuration'
            }
        }

        It 'retains the pinned version failure for a mismatch or nonzero native exit' {
            $original = (Get-Command Invoke-RSClangFormatCapture).ScriptBlock
            try {
                Mock Invoke-RSClangFormatCapture { [pscustomobject]@{ExitCode=0;StdOut=@('clang-format version 21.0.0');ErrorText=''} }
                { Assert-RSClangFormatVersion -Path 'owned-formatter.exe' -Version '22.1.3' } | Should Throw 'does not report clang-format 22.1.3'
                Mock Invoke-RSClangFormatCapture { [pscustomobject]@{ExitCode=17;StdOut=@('clang-format version 22.1.3');ErrorText=''} }
                { Assert-RSClangFormatVersion -Path 'owned-formatter.exe' -Version '22.1.3' } | Should Throw 'does not report clang-format 22.1.3'
                Assert-MockCalled Invoke-RSClangFormatCapture -Times 2 -Exactly -Scope It -ParameterFilter { $Arguments.Count -eq 1 -and $Arguments[0] -ceq '--version' }
            } finally {
                # Pester3 can retain a module-local mock after its It finishes.
                Set-Item -LiteralPath Function:script:Invoke-RSClangFormatCapture -Value $original
            }
        }

        It 'captures native formatter metadata and admitted payload without a console or lost stderr' {
            $fixture = New-RSTestFormatterFixture
            $names = @('RS_FORMATTER_VERSION','RS_FORMATTER_EXIT','RS_FORMATTER_STDERR_COUNT','RS_FORMATTER_STDOUT','RS_FORMATTER_LOG')
            $saved = @{}
            foreach ($name in $names) { $saved[$name] = [Environment]::GetEnvironmentVariable($name, 'Process') }
            $savedConsole = [Console]::OutputEncoding
            $savedInput = $global:OutputEncoding
            $cwd = (Get-Location).Path
            try {
                $env:RS_FORMATTER_VERSION = '22.1.3'; $env:RS_FORMATTER_EXIT = '0'
                $env:RS_FORMATTER_STDERR_COUNT = '0'; $env:RS_FORMATTER_STDOUT = $null
                $env:RS_FORMATTER_LOG = $fixture.LogPath
                $version = Invoke-RSClangFormatCapture -FilePath $fixture.Executable -Arguments @('--version') -WorkingDirectory $fixture.Root
                $version.ExitCode | Should Be 0
                ($version.StdOut -join ',') | Should Be 'clang-format version 22.1.3'
                $version.ErrorText | Should Be ''
                $global:LASTEXITCODE | Should Be 0
                $source = Join-Path $fixture.Root 'owned source with spaces.cpp'
                [IO.File]::WriteAllText($source, 'owned fixture; never formatted')
                $sourceHash = (Get-FileHash -LiteralPath $source).Hash
                $payload = Invoke-RSClangFormatCapture -FilePath $fixture.Executable -Arguments @('-i','--style=file',$source) -WorkingDirectory $fixture.Root
                $payload.ExitCode | Should Be 0
                (Get-FileHash -LiteralPath $source).Hash | Should Be $sourceHash
                $arguments = @(Get-Content -LiteralPath $fixture.LogPath | ForEach-Object { ,($_ | ConvertFrom-Json -NoEnumerate) })
                $arguments.Count | Should Be 1
                ($arguments[0] -join '|') | Should Be ('-i|--style=file|' + $source)
                $env:RS_FORMATTER_STDERR_COUNT = '131072'; $env:RS_FORMATTER_STDOUT = 'separate-stdout'
                $env:RS_FORMATTER_EXIT = '19'
                $failed = Invoke-RSClangFormatCapture -FilePath $fixture.Executable -Arguments @('--version') -WorkingDirectory $fixture.Root
                $failed.ExitCode | Should Be 19
                $global:LASTEXITCODE | Should Be 19
                ($failed.ErrorText.Length -ge 131072) | Should Be $true
                ($failed.StdOut -contains 'separate-stdout') | Should Be $true
                $audit = @(Get-Content -LiteralPath $fixture.AuditPath | ForEach-Object { $_ | ConvertFrom-Json })
                $audit.Count | Should Be 3
                foreach ($row in $audit) {
                    $row.consolePresent | Should Be $false
                    $row.consoleVisible | Should Be $false
                    $row.workingDirectory | Should Be $fixture.Root
                }
                [object]::ReferenceEquals([Console]::OutputEncoding, $savedConsole) | Should Be $true
                [object]::ReferenceEquals($global:OutputEncoding, $savedInput) | Should Be $true
                (Get-Location).Path | Should Be $cwd
            } finally {
                foreach ($name in $names) { [Environment]::SetEnvironmentVariable($name, $saved[$name], 'Process') }
            }
        }

        It 'runs the same transport from its bootstrap module without a restored engine' {
            $fixture = New-RSTestFormatterFixture
            $copyRoot = Join-Path $fixture.Root 'engine-free'
            $build = Join-Path $copyRoot 'Tools/Modules/Build'
            [void][IO.Directory]::CreateDirectory($build)
            $owner = Split-Path (Get-Module ClangFormat).Path
            foreach ($name in @('ClangFormat.psm1','ArtifactOperationLock.psm1')) {
                [IO.File]::Copy((Join-Path $owner $name), (Join-Path $build $name))
            }
            (Test-Path (Join-Path $copyRoot '.ci-engine')) | Should Be $false
            $script = Join-Path $fixture.Root 'engine-free-probe.ps1'
            @'
param($ModulePath,$Executable,$WorkingDirectory)
$ErrorActionPreference='Stop'
Import-Module $ModulePath -ErrorAction Stop
$env:RS_FORMATTER_VERSION='22.1.3';$env:RS_FORMATTER_EXIT='0'
$env:RS_FORMATTER_STDERR_COUNT='0';$env:RS_FORMATTER_STDOUT=$null
$result=Invoke-RSClangFormatCapture -FilePath $Executable -Arguments @('--version') -WorkingDirectory $WorkingDirectory
if($result.ExitCode -ne 0 -or ($result.StdOut -join ',') -cne 'clang-format version 22.1.3'){throw 'Engine-free formatter capture failed.'}
[pscustomobject]@{exit=$result.ExitCode;enginePresent=(Test-Path (Join-Path (Split-Path (Split-Path (Split-Path (Split-Path $ModulePath)))) '.ci-engine'))}|ConvertTo-Json -Compress
'@ | Set-Content -LiteralPath $script -Encoding utf8NoBOM
            $result = Invoke-RSTestFixtureNative -FilePath (Join-Path $PSHOME 'pwsh.exe') -Arguments @('-NoLogo','-NoProfile','-NonInteractive','-File',$script,(Join-Path $build 'ClangFormat.psm1'),$fixture.Executable,$fixture.Root) -WorkingDirectory $fixture.Root
            $global:LASTEXITCODE | Should Be 0
            $proof = ($result -join [Environment]::NewLine) | ConvertFrom-Json
            $proof.exit | Should Be 0
            $proof.enginePresent | Should Be $false
        }
    }
}
