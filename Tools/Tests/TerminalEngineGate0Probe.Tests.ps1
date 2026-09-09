Set-StrictMode -Version Latest

Describe 'Candidate-neutral terminal-engine Gate-0 ABI and fuzz probe' {
    BeforeAll {
        $script:RepositoryRoot = [IO.Path]::GetFullPath(
            (Join-Path $PSScriptRoot '..\..'))
        $script:Source = Join-Path `
            $script:RepositoryRoot `
            'Tools\TerminalEngine\TerminalEngineGate0Probe.cpp'
        $script:Project = Join-Path `
            $script:RepositoryRoot `
            'Tools\TerminalEngine\TerminalEngineGate0Probe.vcxproj'
        $script:Layout = Join-Path `
            $script:RepositoryRoot `
            'Tools\TerminalEngine\TerminalEngineGate0Layout.h'
        $script:Wrapper = Join-Path `
            $script:RepositoryRoot `
            'Tools\TerminalEngine\Invoke-TerminalEngineGate0Probe.ps1'
        $script:Contract = Join-Path `
            $script:RepositoryRoot `
            'Tools\TerminalEngine\TerminalEngineGate0AdapterContract.md'
        $script:TestAdapterSource = Join-Path `
            $script:RepositoryRoot `
            'Tools\TerminalEngine\TerminalEngineGate0ContractTestAdapter.cpp'
        $script:TestAdapterProject = Join-Path `
            $script:RepositoryRoot `
            'Tools\TerminalEngine\TerminalEngineGate0ContractTestAdapter.vcxproj'
    }

    It 'builds the probe with every required evidence lane' {
        $project = Get-Content -LiteralPath $script:Project -Raw
        $wrapper = Get-Content -LiteralPath $script:Wrapper -Raw
        $project | Should Match 'Release\|x64'
        $project | Should Match 'ASan Debug\|x64'
        $project | Should Match 'Release\|ARM64'
        $project | Should Match 'TreatWarningAsError'
        $project | Should Match 'EnableASAN'
        $project | Should Match 'Crypt32\.lib'
        $project | Should Match '<RSUseVcpkgManifest>false'
        $project | Should Match '<VcpkgEnabled>false'
        $project | Should Match '<VcpkgEnableManifest>false'
        $project | Should Match (
            '<AdditionalIncludeDirectories>' +
            '\$\(HostHeaderIncludeRoot\);')
        $project | Should Match 'wil\\resource\.h'
        $project | Should Match 'TerminalEngineGate0Layout\.h'
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

    It 'covers every public enum and structure in the ABI identity' {
        $source = Get-Content -LiteralPath $script:Layout -Raw
        $enums = @(
            'RsTerminalGate0Result',
            'RsTerminalGate0Architecture',
            'RsTerminalGate0CallingConvention',
            'RsTerminalGate0ResourceKind',
            'RsTerminalGate0Capability',
            'RsTerminalGate0ActionFlag'
        )
        foreach ($name in $enums) {
            $source | Should Match (
                'AppendEnum<' + [regex]::Escape($name) + '>')
        }

        $structs = @(
            'RsTerminalGate0String',
            'RsTerminalGate0IdentityV1',
            'RsTerminalGate0ResourceLimitsV1',
            'RsTerminalGate0ResourceCallbacksV1',
            'RsTerminalGate0SessionConfigV1',
            'RsTerminalGate0ActionV1',
            'RsTerminalGate0ActionObservationV1',
            'RsTerminalGate0AdapterV1'
        )
        foreach ($name in $structs) {
            $source | Should Match (
                'AppendStruct<' + [regex]::Escape($name) + '>')
        }
        $source | Should Match 'offsetof\(type,\s*field\)'
        $source | Should Match 'alignof\(decltype'
        (Get-Content -LiteralPath $script:Source -Raw) |
            Should Match 'layout-v1-'
        $source | Should Match 'BuildLayoutJcs'
    }

    It 'includes every public C function-pointer layout' {
        $source = Get-Content -LiteralPath $script:Layout -Raw
        $functionPointers = @(
            'RsTerminalGate0ReserveResourceFn',
            'RsTerminalGate0ReleaseResourceFn',
            'RsTerminalGate0SessionCreateFn',
            'RsTerminalGate0SessionWriteFn',
            'RsTerminalGate0SessionActionFn',
            'RsTerminalGate0SessionStateJcsFn',
            'RsTerminalGate0SessionDestroyFn',
            'RsTerminalGate0QueryAdapterFn',
            'RsTerminalGate0QueryLayoutJcsFn'
        )
        foreach ($name in $functionPointers) {
            $source | Should Match (
                'AppendFunctionPointer<' +
                [regex]::Escape($name) +
                '>')
        }
    }

    It 'keeps the local parser robustness corpus deterministic and bounded' {
        $source = Get-Content -LiteralPath $script:Source -Raw
        $source | Should Match 'kMaximumRandomCases = 64'
        $source | Should Match 'kMaximumCaseBytes = 1024'
        $source | Should Match 'kMaximumDenialOrdinals = 128'
        $source | Should Match 'corpus\.push_back\(\{\.bytes = \{\}, \.chunks = \{0\}\}\)'
        $source | Should Match 'const std::vector<uint8_t> malformed'
        $source | Should Match 'std::vector<size_t>\(malformed\.size\(\), 1\)'
        $source | Should Match 'DeterministicGenerator generator\(arguments\.seed\)'
        $source | Should Match 'fixedStateHashes\[0\] == fixedStateHashes\[1\]'
        $source | Should Match 'SameTrace\(first, second\)'
    }

    It 'sweeps denial ordinals and proves destroy quiet point and unload' {
        $source = Get-Content -LiteralPath $script:Source -Raw
        $source | Should Match 'ordinal <= summary\.denialSweepCount'
        $source | Should Match 'first\.denialObserved'
        $source | Should Match 'adapter\.session_destroy\(session\)'
        $source | Should Match 'ledger\.Close\(\)'
        $source | Should Match 'ledger\.IsQuiet\(\)'
        $source | Should Match 'module\.reset\(\)'
        $source | Should Match 'GetModuleHandleW'
        $source | Should Not Match 'FreeLibrary\('
        $source | Should Not Match 'catch\s*\(\s*\.\.\.\s*\)'
    }

    It 'emits bounded content-free fuzz evidence' {
        $source = Get-Content -LiteralPath $script:Source -Raw
        $source | Should Match 'adapterIdentitySha256'
        $source | Should Match 'adapterLayoutSha256'
        $source | Should Match 'resourceTraceSha256'
        $source | Should Match 'resultTraceSha256'
        $source | Should Match 'stateTraceSha256'
        $source | Should Match 'summary\.unloadCount == summary\.runCount'
        $source | Should Not Match '"inputBytes"'
        $source | Should Not Match '"terminalState"'
        $source | Should Not Match '"adapterPath"'
    }

    It 'compares adapter-TU layout bytes and every frozen identity exactly' {
        $source = Get-Content -LiteralPath $script:Source -Raw
        $layout = Get-Content -LiteralPath $script:Layout -Raw
        $source | Should Match 'QueryAdapterLayoutJcs'
        $source | Should Match 'kRsTerminalGate0QueryLayoutJcsSymbol'
        $source | Should Match (
            'adapterLayout\s*==\s*arguments\.probeLayoutJcs')
        $source | Should Match (
            'adapter\.identity\.candidate_id\)\s*==\s*' +
            'arguments\.expectedCandidateId')
        $source | Should Match (
            'adapter\.identity\.upstream_pin\)\s*==\s*' +
            'arguments\.expectedUpstreamPin')
        $source | Should Match (
            'adapter\.identity\.source_sha256\)\s*==\s*' +
            'arguments\.expectedSourceSha256')
        $source | Should Match (
            'adapter\.identity\.adapter_header_sha256\)\s*==\s*' +
            'arguments\.expectedAdapterHeaderSha256')
        $source | Should Match (
            'adapter\.identity\.build_identity\)\s*==\s*' +
            'arguments\.expectedBuildIdentity')
        $source | Should Match (
            'adapter\.identity\.capabilities\s*==\s*' +
            'arguments\.expectedCapabilities')
        $layout | Should Match 'QueryLayoutJcs'
        $layout | Should Match 'kRsTerminalGate0LayoutQueryVersion'
        $layout | Should Not Match 'RS_GATE0_LAYOUT_IDENTITY'
    }

    It 'validates the canonical layout digest and fuzz result in the wrapper' {
        $errors = $null
        [void][Management.Automation.Language.Parser]::ParseFile(
            $script:Wrapper,
            [ref]$null,
            [ref]$errors)
        @($errors).Count | Should Be 0
        $wrapper = Get-Content -LiteralPath $script:Wrapper -Raw
        $wrapper | Should Match 'Get-EmbeddedLayoutJcs'
        $wrapper | Should Match 'layout-v1-\$actualSha256'
        $wrapper | Should Match 'layoutJcs\.enums\)\.Count -ne 6'
        $wrapper | Should Match 'layoutJcs\.functionPointers\)\.Count -ne 9'
        $wrapper | Should Match 'layoutJcs\.structs\)\.Count -ne 8'
        $wrapper | Should Match 'runCount -ne'
        $wrapper | Should Match 'unloadCount'
        $wrapper | Should Match 'denialSweepCount'
    }

    It 'runs the frozen C7 stream through real adapter sessions and validates the exact proof envelope' {
        $source = Get-Content -LiteralPath $script:Source -Raw
        $project = Get-Content -LiteralPath $script:Project -Raw
        $wrapper = Get-Content -LiteralPath $script:Wrapper -Raw
        $project | Should Match (
            'Generated\\TerminalEngineFixtures\.v1\.h')
        $wrapper | Should Match (
            "ValidateSet\('layout', 'fuzz', 'c7'\)")
        $wrapper | Should Match (
            'Get-RSTerminalEngineC7FixedStreamBytes')
        $wrapper | Should Match "'--c7-stream-hex'"
        $wrapper | Should Match (
            '\[Convert\]::ToHexString\(\$fixedStream\)' +
            '\.ToLowerInvariant\(\)')
        $wrapper | Should Match (
            "expectedSchedules = \[string\[\]\]@\(")
        $wrapper | Should Match 'sessionCount -ne 4'
        $wrapper | Should Match 'quietCount -ne 4'
        $wrapper | Should Match 'unloadCount -ne 1'
        $source | Should Match (
            '#include "Generated/TerminalEngineFixtures\.v1\.h"')
        $source | Should Match 'RunC7Session\('
        $source | Should Match 'adapter\.session_create\('
        $source | Should Match 'adapter\.session_write\('
        $source | Should Match 'adapter\.session_action\('
        $source | Should Match 'adapter\.session_state_jcs\('
        $source | Should Match 'adapter\.session_destroy\('
        $source | Should Match '"write-response-order"'
        $source | Should Match '"admit-external-events"'
        $source | Should Match 'summary\.quietCount == 4'
        $source | Should Match 'summary\.unloaded'
        $source | Should Match (
            'requiredOrderingFlags')
    }

    It 'makes the contract adapter derive layout in its own translation unit' {
        $source = Get-Content `
            -LiteralPath $script:TestAdapterSource `
            -Raw
        $project = Get-Content `
            -LiteralPath $script:TestAdapterProject `
            -Raw
        $source | Should Match 'TerminalEngineGate0Layout\.h'
        $source | Should Match 'rs_terminal_gate0_query_layout_jcs'
        $source | Should Match 'QueryLayoutJcs'
        $source | Should Not Match 'RS_GATE0_LAYOUT_IDENTITY'
        $project | Should Match 'TerminalEngineGate0Layout\.h'
        $project | Should Not Match 'Gate0LayoutIdentity'
    }

    It 'documents layout derivation and public fuzz-seam failure rules' {
        $contract = Get-Content -LiteralPath $script:Contract -Raw
        $contract | Should Match 'layout-v1-<sha256-of-layout-JCS>'
        $contract | Should Match 'rs_terminal_gate0_query_layout_jcs'
        $contract | Should Not Match 'RS_GATE0_LAYOUT_IDENTITY'
        $contract | Should Match 'fixed\s+empty, malformed UTF-8/VT'
        $contract | Should Match 'allocation-denial ordinal'
        $contract | Should Match 'module remaining'
    }
}
