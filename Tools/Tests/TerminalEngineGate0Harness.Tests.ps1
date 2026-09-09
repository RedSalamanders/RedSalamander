Set-StrictMode -Version Latest

Describe 'Candidate-neutral terminal-engine Gate-0 harness' {
    BeforeAll {
        $script:RepositoryRoot = [IO.Path]::GetFullPath(
            (Join-Path $PSScriptRoot '..\..'))
        $script:HarnessSource = Join-Path `
            $script:RepositoryRoot `
            'Tools\TerminalEngine\TerminalEngineGate0Harness.cpp'
        $script:AdapterHeader = Join-Path `
            $script:RepositoryRoot `
            'Tools\TerminalEngine\TerminalEngineGate0Adapter.h'
        $script:AdapterContract = Join-Path `
            $script:RepositoryRoot `
            'Tools\TerminalEngine\TerminalEngineGate0AdapterContract.md'
        $script:Project = Join-Path `
            $script:RepositoryRoot `
            'Tools\TerminalEngine\TerminalEngineGate0Harness.vcxproj'
        $script:Wrapper = Join-Path `
            $script:RepositoryRoot `
            'Tools\TerminalEngine\Invoke-TerminalEngineGate0Harness.ps1'
        $script:ContractTestAdapterSource = Join-Path `
            $script:RepositoryRoot `
            'Tools\TerminalEngine\TerminalEngineGate0ContractTestAdapter.cpp'
        $script:ContractTestAdapterProject = Join-Path `
            $script:RepositoryRoot `
            'Tools\TerminalEngine\TerminalEngineGate0ContractTestAdapter.vcxproj'
        $script:LifecycleModule = Join-Path `
            $script:RepositoryRoot `
            'Tools\TerminalEngine\TerminalEngineLifecycle.psm1'
    }

    It 'keeps the reviewed neutral ABI frozen' {
        (Get-FileHash `
            -LiteralPath $script:AdapterHeader `
            -Algorithm SHA256).Hash.ToLowerInvariant() |
            Should Be `
                '5aed4f300a5a4aea090f50bbed743bd34824f18a9d2acc64d698a9b8561981bc'
        $header = Get-Content -LiteralPath $script:AdapterHeader -Raw
        $header | Should Match 'rs_terminal_gate0_query_adapter'
        $header | Should Match 'RsTerminalGate0SessionActionFn'
        $header | Should Match 'RsTerminalGate0SessionStateJcsFn'
        $header | Should Match 'fixture_id'
        $header | Should Match 'schedule_id'
    }

    It 'requires every frozen row action and exact check result' {
        $source = Get-Content -LiteralPath $script:HarnessSource -Raw
        $source | Should Match 'scheduleCount == 43'
        $source | Should Match 'fixture\.actions\.front\(\)\.actionId'
        $source | Should Match 'adapter\.session_action'
        $source | Should Match 'actualState == expectedState->expectedValue'
        $source | Should Match 'actualResource == expectedResource->expectedValue'
        $source | Should Match 'fixture\.checks\.size\(\) != 2'
        $source | Should Match 'callback-after-destroy'
    }

    It 'freezes all nineteen non-feed action semantics' {
        $contract = Get-Content -LiteralPath $script:AdapterContract -Raw
        $actions = @(
            'admit-external-events',
            'encode-focus-gained-lost',
            'encode-mouse-matrix',
            'encode-paste-disabled',
            'encode-paste-enabled',
            'probe-generic-escape-limit',
            'probe-kitty-pixel-limit',
            'probe-kitty-replacement-ledger',
            'probe-osc52-limits',
            'probe-osc8-limits',
            'probe-pending-effect-ledger',
            'probe-terminal-text-ledger',
            'probe-title-cwd-limits',
            'resize-sequence',
            'select-logical-range',
            'snapshot-copy-fail-retry',
            'snapshot-end-mutate-probe',
            'snapshot-static-kitty',
            'teardown-sequence'
        )
        foreach ($action in $actions) {
            $contract | Should Match ([regex]::Escape("``$action``"))
        }
    }

    It 'supports every lane only with an explicit frozen host-header root' {
        $project = Get-Content -LiteralPath $script:Project -Raw
        $wrapper = Get-Content -LiteralPath $script:Wrapper -Raw
        $project | Should Match 'Release\|x64'
        $project | Should Match 'ASan Debug\|x64'
        $project | Should Match 'Release\|ARM64'
        $project | Should Match 'TreatWarningAsError'
        $project | Should Match 'EnableASAN'
        $project | Should Match '<RSUseVcpkgManifest>false'
        $project | Should Match '<VcpkgEnabled>false'
        $project | Should Match '<VcpkgEnableManifest>false'
        $project | Should Match (
            '<AdditionalIncludeDirectories>' +
            '\$\(HostHeaderIncludeRoot\);' +
            '\$\(FixtureIncludeRoot\);')
        $project | Should Match (
            "HostHeaderIncludeRoot\)'==''")
        $project | Should Match 'wil\\resource\.h'
        $wrapper | Should Match (
            '\[string\]\$HostHeaderIncludeRoot')
        $wrapper | Should Match (
            '\[string\]\$ExpectedHostHeaderTreeSha256')
        $wrapper | Should Match (
            'Resolve-RSTerminalGate0HostHeaderRoot')
        ([regex]::Matches(
                $wrapper,
                'Resolve-RSTerminalGate0HostHeaderRoot')).Count |
            Should Be 3
        $wrapper | Should Match (
            '/p:RSUseVcpkgManifest=false')
        $wrapper | Should Match (
            '/p:VcpkgEnabled=false')
        $wrapper | Should Match (
            '/p:VcpkgEnableManifest=false')
        $wrapper | Should Match (
            '/p:HostHeaderIncludeRoot=')
    }

    It 'cannot fall back to a mutated ambient repo WIL tree' {
        Import-Module $script:LifecycleModule -Force
        $explicitRoot = Join-Path $TestDrive 'frozen-host\include'
        $ambientRoot = Join-Path $TestDrive (
            'repo\.build\vcpkg_installed\x64-windows\include')
        [void][IO.Directory]::CreateDirectory(
            (Join-Path $explicitRoot 'wil'))
        [void][IO.Directory]::CreateDirectory(
            (Join-Path $ambientRoot 'wil'))
        [IO.File]::WriteAllText(
            (Join-Path $explicitRoot 'wil\resource.h'),
            "frozen`n",
            [Text.UTF8Encoding]::new($false))
        $ambientHeader = Join-Path $ambientRoot 'wil\resource.h'
        [IO.File]::WriteAllText(
            $ambientHeader,
            "ambient-original`n",
            [Text.UTF8Encoding]::new($false))
        $expected = [string](
            Get-RSDirectoryIdentity -Path $explicitRoot).sha256

        [IO.File]::WriteAllText(
            $ambientHeader,
            "ambient-mutated`n",
            [Text.UTF8Encoding]::new($false))
        $resolved = Resolve-RSTerminalGate0HostHeaderRoot `
            -HostHeaderIncludeRoot $explicitRoot `
            -ExpectedTreeSha256 $expected
        $resolved.IncludeRoot |
            Should Be ([IO.Path]::GetFullPath($explicitRoot))
        $resolved.TreeSha256 | Should Be $expected

        [IO.File]::Delete(
            (Join-Path $explicitRoot 'wil\resource.h'))
        $threw = $false
        try {
            $null = Resolve-RSTerminalGate0HostHeaderRoot `
                -HostHeaderIncludeRoot $explicitRoot `
                -ExpectedTreeSha256 $expected
        }
        catch [IO.InvalidDataException] {
            $threw = $true
        }
        $threw | Should Be $true
    }

    It 'rejects a host-header root reached through an ancestor junction' {
        Import-Module $script:LifecycleModule -Force
        $realRoot = Join-Path $TestDrive 'real-host'
        $realInclude = Join-Path $realRoot 'include'
        [void][IO.Directory]::CreateDirectory(
            (Join-Path $realInclude 'wil'))
        [IO.File]::WriteAllText(
            (Join-Path $realInclude 'wil\resource.h'),
            "frozen`n",
            [Text.UTF8Encoding]::new($false))
        $expected = [string](
            Get-RSDirectoryIdentity -Path $realInclude).sha256
        $junction = Join-Path $TestDrive 'host-junction'
        $null = New-Item `
            -ItemType Junction `
            -Path $junction `
            -Target $realRoot

        $threw = $false
        try {
            $null = Resolve-RSTerminalGate0HostHeaderRoot `
                -HostHeaderIncludeRoot (Join-Path $junction 'include') `
                -ExpectedTreeSha256 $expected
        }
        catch {
            $threw = $true
        }
        $threw | Should Be $true
    }

    It 'validates successful output cardinality and uniqueness in the wrapper' {
        $errors = $null
        [void][Management.Automation.Language.Parser]::ParseFile(
            $script:Wrapper,
            [ref]$null,
            [ref]$errors)
        @($errors).Count | Should Be 0
        $wrapper = Get-Content -LiteralPath $script:Wrapper -Raw
        $wrapper | Should Match '\$rows\.Count -ne 43'
        $wrapper | Should Match 'ContainsKey\(\$key\)'
        $wrapper | Should Match 'expectedStateSha256'
        $wrapper | Should Match 'expectedResourceTraceSha256'
        $wrapper | Should Match 'limitations'
        $wrapper | Should Match '\$actual\.unit -cne \$expected\.Unit'
        $wrapper | Should Match '\$samples\.Count -ne \$expected\.Samples'
        $wrapper | Should Match '\$sample -isnot \[string\]'
        $wrapper | Should Match '\^\(\?:0\|\[1-9\]\[0-9\]\{0,19\}\)\$'
        $wrapper | Should Match '\[uint64\]::TryParse'
    }

    It 'keeps the oracle-reading contract adapter test-only' {
        $source = Get-Content `
            -LiteralPath $script:ContractTestAdapterSource `
            -Raw
        $source | Should Match 'Harness mechanics test double only'
        $source | Should Match 'MUST NOT be used as candidate or promotion'
        $source | Should Match 'evidence, shipped, or linked into the product'
        $source | Should Match 'contract-test-adapter'
        $source | Should Match 'TerminalEngineFixtures\.v1\.h'
        $source | Should Match 'RS_TERMINAL_GATE0_CONTRACT_TEST_MODE'
        $source | Should Match 'state-mismatch'
        $source | Should Match 'action-unsupported'
        $source | Should Match 'limit-evidence-missing'
        $source | Should Match 'wrong-capabilities'

        $project = Get-Content `
            -LiteralPath $script:ContractTestAdapterProject `
            -Raw
        $project | Should Match 'DynamicLibrary'
        $project | Should Match 'Release\|x64'
        $project | Should Match 'ASan Debug\|x64'
        $project | Should Match 'Release\|ARM64'
        $project | Should Match 'TreatWarningAsError'
    }

    It 'requires the complete candidate-neutral performance contract' {
        $source = Get-Content -LiteralPath $script:HarnessSource -Raw
        $contract = Get-Content -LiteralPath $script:AdapterContract -Raw
        $metrics = @(
            'create-destroy-us',
            'parse-8mib-us',
            'full-snapshot-us',
            'dirty-snapshot-us',
            'resize-reflow-us',
            'kitty-1024-rgba-us',
            'eight-session-storm-us',
            'eight-session-retained-bytes',
            'added-release-package-bytes',
            'crossed-abi-elements',
            'non-system-runtime-dependencies'
        )
        foreach ($metric in $metrics) {
            $source | Should Match ([regex]::Escape("`"$metric`""))
        }
        $contract | Should Match 'benchmark-full-snapshot'
        $contract | Should Match 'benchmark-dirty-snapshot'
        $contract | Should Match 'benchmark-resize-reflow'
        $contract | Should Match '80x50,120x40'
        $contract | Should Match 'benchmark-kitty-1024-rgba'
    }

    It 'implements the frozen measurement scenarios without weaker substitutes' {
        $source = Get-Content -LiteralPath $script:HarnessSource -Raw
        $contract = Get-Content -LiteralPath $script:AdapterContract -Raw

        $source | Should Match 'fractionalRemainder'
        $source | Should Match 'greaterThanHalf'
        $source | Should Match 'exactHalf'
        $source | Should Match 'chunks\[scheduleIndex\+\+ % chunks\.size\(\)\]'
        $source | Should Match 'constexpr size_t chunkBytes = 4096U'
        $source | Should Match 'for \(const auto& session : stormSessions\)'
        $source | Should Match 'CurrentCpuBytes\(\)'
        $source | Should Match 'retained\.warmupCount = 1'
        $source | Should Match 'retained\.sampleCount = 5'
        $source | Should Match "index < 10'000U"
        $source | Should Match '\\x1b\[1;1HA\\x1b\[2;1HB\\x1b\[3;3H'
        $source | Should Match '"fibonacci"'
        $source | Should Match '80x50,120x40'
        $source | Should Not Match '80x24,20x24,100x24'

        $contract | Should Match 'round-robin 4,096-byte chunks'
        $contract | Should Match 'CPU-only retained ownership'
        $contract | Should Match 'snapshot, and delete'
        $contract | Should Match '1 warmup, 5 samples'
        $contract | Should Match 'terminal-engine-contribution-zip-v1'
        $contract | Should Match 'Raw DLL plus notice sizes are not a valid substitute'
    }
}
