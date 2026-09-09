$script:RepoRoot = (Resolve-Path (Join-Path $PSScriptRoot '..\..')).Path
$script:InventoryModulePath = Join-Path $script:RepoRoot 'Tools\Modules\Tooling\ToolInventory.psm1'

function Get-RSToolingGovernanceManifest {
    return Get-Content -LiteralPath (Join-Path $script:RepoRoot 'Tools\tool-inventory.json') -Raw |
        ConvertFrom-Json
}

function Get-RSPowerShellAst {
    param([Parameter(Mandatory = $true)][string]$Path)

    $tokens = $null
    $errors = $null
    $ast = [System.Management.Automation.Language.Parser]::ParseFile(
        $Path,
        [ref]$tokens,
        [ref]$errors)
    if (@($errors).Count -ne 0) {
        throw "PowerShell parse failure in '$Path': $(@($errors.Message) -join '; ')"
    }
    return $ast
}

function Assert-RSNoToolingGovernanceFailures {
    param(
        [Parameter(Mandatory = $true)][AllowEmptyCollection()][object[]]$Failures,
        [Parameter(Mandatory = $true)][string]$Contract
    )

    if (@($Failures).Count -ne 0) {
        throw "$Contract`n$(@($Failures) -join "`n")"
    }
}

function Get-RSMarkdownTableFirstColumn {
    param(
        [Parameter(Mandatory = $true)][AllowEmptyString()][string[]]$Lines,
        [Parameter(Mandatory = $true)][string]$Heading
    )

    $headingIndex = [Array]::IndexOf($Lines, $Heading)
    if ($headingIndex -lt 0) {
        throw "Markdown heading '$Heading' was not found."
    }

    $values = [System.Collections.Generic.List[string]]::new()
    $insideTable = $false
    for ($index = $headingIndex + 1; $index -lt $Lines.Count; $index++) {
        $line = $Lines[$index]
        if ($line -match '^\s*\|') {
            $insideTable = $true
            if ($line -match '^\s*\|\s*`(?<Value>[^`]+)`\s*\|' -and
                $Matches.Value -notin @('File', 'Module')) {
                $values.Add($Matches.Value)
            }
            continue
        }
        if ($insideTable) {
            break
        }
    }

    return @($values)
}

Describe 'Repository tooling governance' {
    BeforeAll {
        Import-Module $script:InventoryModulePath -Force
    }

    It 'keeps the checked-in inventory reconciled with the Tools tree' {
        $inventory = Get-RSToolInventory -RepositoryRoot $script:RepoRoot

        if ($inventory.Summary.BlockingFindings -ne 0) {
            throw "Tool inventory has $($inventory.Summary.BlockingFindings) finding(s): $(@($inventory.Findings | ForEach-Object { "$($_.Code): $($_.Path)" }) -join '; ')"
        }
        if ($inventory.Summary.ClassifiedFiles -ne $inventory.Summary.DiscoveredFiles) {
            throw "Tool inventory classified $($inventory.Summary.ClassifiedFiles) of $($inventory.Summary.DiscoveredFiles) discovered files."
        }
    }

    It 'keeps active private implementation modules below Tools Modules' {
        $manifest = Get-RSToolingGovernanceManifest
        $failures = @($manifest.entries | Where-Object {
                $_.kind -eq 'module' -and
                $_.lifecycle -eq 'active' -and
                [string]$_.path -match '^Tools/[^/]+$'
            } | ForEach-Object {
                "$($_.path): active private implementation modules belong below Tools/Modules and use .psm1."
            })

        Assert-RSNoToolingGovernanceFailures -Failures $failures `
            -Contract 'Root-level active function libraries are forbidden.'
    }

    It 'limits internal root compatibility loaders to explicit module replacements' {
        $manifest = Get-RSToolingGovernanceManifest
        $failures = @($manifest.entries | Where-Object {
                $_.visibility -eq 'internal' -and
                $_.lifecycle -eq 'compatibility' -and
                [string]$_.path -match '^Tools/[^/]+\.ps1$'
            } | ForEach-Object {
                if ([string]::IsNullOrWhiteSpace([string]$_.replacement) -or
                    [string]$_.replacement -notmatch '^Tools/Modules/.+\.psm1$') {
                    "$($_.path): an internal root compatibility loader must name its canonical Tools/Modules .psm1 replacement."
                }
            })

        Assert-RSNoToolingGovernanceFailures -Failures $failures `
            -Contract 'Internal compatibility loaders require explicit canonical replacements.'
    }

    It 'documents every public command and compatibility loader completely' {
        $manifest = Get-RSToolingGovernanceManifest
        $entries = @($manifest.entries | Where-Object {
                [string]$_.path -like '*.ps1' -and
                ($_.visibility -eq 'public' -or $_.lifecycle -eq 'compatibility')
            })
        $failures = [System.Collections.Generic.List[string]]::new()
        foreach ($entry in $entries) {
            $fullPath = Join-Path $script:RepoRoot ([string]$entry.path).Replace('/', '\')
            $source = Get-Content -LiteralPath $fullPath -Raw
            $ast = Get-RSPowerShellAst -Path $fullPath
            $help = Get-Help $fullPath -ErrorAction Stop
            if ([string]::IsNullOrWhiteSpace([string]$help.Synopsis) -or
                [string]$help.Synopsis -eq [IO.Path]::GetFileName([string]$entry.path)) {
                $failures.Add("$($entry.path): comment-based help is not discoverable through Get-Help.")
            }

            foreach ($heading in @('SYNOPSIS', 'DESCRIPTION', 'OUTPUTS', 'NOTES', 'EXAMPLE')) {
                if ($source -notmatch "(?im)^\s*\.$heading(?:\s|$)") {
                    $failures.Add("$($entry.path): missing .$heading comment-based help.")
                }
            }

            $scriptParameters = if ($null -eq $ast.ParamBlock) { @() } else { @($ast.ParamBlock.Parameters) }
            foreach ($parameter in $scriptParameters) {
                $parameterName = $parameter.Name.VariablePath.UserPath
                $isHidden = $parameter.Extent.Text -match '(?i)DontShow'
                if (-not $isHidden -and
                    $source -notmatch "(?im)^\s*\.PARAMETER\s+$([regex]::Escape($parameterName))(?:\s|$)") {
                    $failures.Add("$($entry.path): missing .PARAMETER $parameterName help.")
                }
            }
        }

        Assert-RSNoToolingGovernanceFailures -Failures @($failures) `
            -Contract 'Public commands and compatibility loaders require complete comment-based help.'
    }

    It 'requires canonical replacements and explicit retention policy for compatibility surfaces' {
        $manifest = Get-RSToolingGovernanceManifest
        $failures = [System.Collections.Generic.List[string]]::new()
        foreach ($entry in @($manifest.entries | Where-Object {
                    $_.kind -eq 'compatibility' -or
                    $_.lifecycle -eq 'compatibility'
                })) {
            if ([string]::IsNullOrWhiteSpace([string]$entry.supportPolicy)) {
                $failures.Add("$($entry.path): compatibility metadata has no supportPolicy.")
            }

            if ($entry.kind -ne 'compatibility') {
                continue
            }
            $replacement = [string]$entry.replacement
            if ([string]::IsNullOrWhiteSpace($replacement)) {
                $failures.Add("$($entry.path): compatibility metadata has no replacement.")
            }
            elseif ($replacement -notmatch '^Tools/.+' -or
                -not (Test-Path -LiteralPath (Join-Path $script:RepoRoot $replacement.Replace('/', '\')))) {
                $failures.Add("$($entry.path): replacement '$replacement' must be an existing exact repository-relative Tools path.")
            }
        }

        Assert-RSNoToolingGovernanceFailures -Failures @($failures) `
            -Contract 'Compatibility surfaces require an exact replacement and every retained dormant surface requires supportPolicy.'
    }

    It 'keeps every historical file reachable from a reviewed manual command root' {
        $manifest = Get-RSToolingGovernanceManifest
        $reviewedRoots = [string[]]@(
            'Tools/TerminalEngine/Invoke-HistoricalTerminalTests.ps1',
            'Tools/Evaluate-TerminalEngines.ps1',
            'Tools/Test-TerminalEngineRound4Closeout.ps1'
        )
        $entryByPath = @{}
        foreach ($entry in @($manifest.entries)) {
            $entryByPath[[string]$entry.path] = $entry
        }

        $failures = [System.Collections.Generic.List[string]]::new()
        foreach ($entry in @($manifest.entries | Where-Object lifecycle -eq 'historical')) {
            $manualEntryPoints = @($entry.manualEntryPoints)
            if ($manualEntryPoints.Count -eq 0) {
                $failures.Add("$($entry.path): historical entry is orphaned from the reviewed manual commands.")
                continue
            }

            foreach ($manualEntryPoint in $manualEntryPoints) {
                $manualEntryPoint = [string]$manualEntryPoint
                if ($reviewedRoots -cnotcontains $manualEntryPoint) {
                    $failures.Add("$($entry.path): '$manualEntryPoint' is not an exact reviewed manual command path.")
                    continue
                }
                if (-not (Test-Path -LiteralPath (Join-Path $script:RepoRoot $manualEntryPoint.Replace('/', '\')) -PathType Leaf)) {
                    $failures.Add("$($entry.path): reviewed manual command '$manualEntryPoint' does not exist.")
                    continue
                }
                if (-not $entryByPath.ContainsKey($manualEntryPoint)) {
                    $failures.Add("$($entry.path): reviewed manual command '$manualEntryPoint' has no explicit inventory entry.")
                    continue
                }

                $rootEntry = $entryByPath[$manualEntryPoint]
                if ([string]$rootEntry.path -cne $manualEntryPoint -or
                    [string]$rootEntry.lifecycle -cne 'historical' -or
                    [string]$rootEntry.kind -cne 'command') {
                    $failures.Add("$($entry.path): '$manualEntryPoint' must resolve exactly to a historical command entry.")
                }
            }
        }

        foreach ($reviewedRoot in $reviewedRoots) {
            if (-not $entryByPath.ContainsKey($reviewedRoot) -or
                @($entryByPath[$reviewedRoot].manualEntryPoints) -cnotcontains $reviewedRoot) {
                $failures.Add("${reviewedRoot}: reviewed manual roots must retain themselves explicitly.")
            }
        }

        Assert-RSNoToolingGovernanceFailures -Failures @($failures) `
            -Contract 'Historical retention requires a closed graph rooted only in the three reviewed manual commands.'
    }

    It 'keeps active executable tooling covered by at least one current test contract' {
        $manifest = Get-RSToolingGovernanceManifest
        $profiles = Get-Content -LiteralPath (Join-Path $script:RepoRoot 'Tools\TerminalEngine\HistoricalTerminalProfiles.json') -Raw |
            ConvertFrom-Json
        $sealedTests = @($profiles.historicalTests.path | ForEach-Object {
                ([string]$_).Replace('\', '/')
            })
        if ($sealedTests.Count -eq 0 -or
            'Tools/Tests/TerminalDependencyLifecycle.Tests.ps1' -notin $sealedTests) {
            throw 'HistoricalTerminalProfiles.json must expose its reviewed sealed cohort through historicalTests[].path.'
        }
        $failures = @($manifest.entries | Where-Object {
                $_.lifecycle -eq 'active' -and
                $_.kind -in @('command', 'module', 'compatibility')
            } | ForEach-Object {
                if (@($_.tests).Count -eq 0) {
                    "$($_.path): active executable tooling must name at least one current test contract."
                }
                elseif (@($_.tests | Where-Object { $_ -notin $sealedTests }).Count -eq 0) {
                    "$($_.path): active tooling cannot rely only on sealed historical tests ($(@($_.tests) -join ', '))."
                }
            })

        Assert-RSNoToolingGovernanceFailures -Failures $failures `
            -Contract 'Active executable tooling needs a current-profile test contract when tests are named.'
    }

    It 'keeps the README top-level and module maps in exact filesystem parity' {
        $readmeLines = @(Get-Content -LiteralPath (Join-Path $script:RepoRoot 'Tools\README.md'))
        $documentedRoot = @(Get-RSMarkdownTableFirstColumn -Lines $readmeLines -Heading '## Top-level inventory' | Sort-Object -Unique)
        $actualRoot = @(Get-ChildItem -LiteralPath (Join-Path $script:RepoRoot 'Tools') -File |
                ForEach-Object Name | Sort-Object -Unique)
        $documentedModules = @(Get-RSMarkdownTableFirstColumn -Lines $readmeLines -Heading '## Implementation module map' |
                Sort-Object -Unique)
        $actualModules = @(Get-ChildItem -LiteralPath (Join-Path $script:RepoRoot 'Tools\Modules') -Recurse -File -Filter '*.psm1' |
                ForEach-Object {
                    $_.FullName.Substring($script:RepoRoot.Length + 1).Replace('\', '/').Substring('Tools/'.Length)
                } | Sort-Object -Unique)

        $rootDelta = @(Compare-Object -ReferenceObject $actualRoot -DifferenceObject $documentedRoot |
                ForEach-Object { "$($_.SideIndicator) $($_.InputObject)" })
        $moduleDelta = @(Compare-Object -ReferenceObject $actualModules -DifferenceObject $documentedModules |
                ForEach-Object { "$($_.SideIndicator) $($_.InputObject)" })
        $failures = @(
            if ($rootDelta.Count -ne 0) {
                "Top-level README map is $($documentedRoot.Count)/$($actualRoot.Count): $($rootDelta -join '; ')"
            }
            if ($moduleDelta.Count -ne 0) {
                "Module README map is $($documentedModules.Count)/$($actualModules.Count): $($moduleDelta -join '; ')"
            }
        )

        Assert-RSNoToolingGovernanceFailures -Failures $failures `
            -Contract 'README navigation must remain a complete dynamic projection of root files and implementation modules.'
    }

    It 'keeps current Terminal helper export and consumer contracts live in the current profile' {
        $contracts = @(
            [pscustomobject]@{
                Module = 'Tools/TerminalEngine/GhosttyRuntimeLifecycle.psm1'
                Exports = @(
                    'Read-RSGhosttyRuntimeLock',
                    'Get-RSGhosttyRuntimeLockDigest',
                    'Get-RSGhosttyRuntimeLockPlatform',
                    'Get-RSGhosttyRuntimeLockDependency',
                    'Resolve-RSGhosttyRuntimeBuildPlan',
                    'Assert-RSGhosttyRuntimeOutputReplaceable',
                    'Invoke-RSGhosttyRuntimeBuild')
                Consumers = @(
                    'Tools/Build-GhosttyTerminalRuntime.ps1',
                    'Tools/TerminalEngine/GhosttyCandidateOverlay.psm1',
                    'Tools/TerminalEngine/GhosttyCandidateReview.psm1',
                    'Tools/TerminalEngine/GhosttyRuntimeStatus.psm1')
            },
            [pscustomobject]@{
                Module = 'Tools/TerminalEngine/GhosttyCandidateOverlay.psm1'
                Exports = @(
                    'Get-RSGhosttyCandidateOverlayAbiReview',
                    'New-RSGhosttyRuntimeCandidateOverlayReview')
                Consumers = @('Tools/New-GhosttyTerminalRuntimeCandidateOverlayReview.ps1')
            },
            [pscustomobject]@{
                Module = 'Tools/TerminalEngine/GhosttyCandidateReview.psm1'
                Exports = @(
                    'Get-RSGhosttyCandidatePathCategory',
                    'Get-RSGhosttyCandidateChangeInventory',
                    'Get-RSGhosttyCandidateAbiReview',
                    'Get-RSGhosttyCandidateSecurityReview',
                    'New-RSGhosttyRuntimeCandidateReview')
                Consumers = @('Tools/New-GhosttyTerminalRuntimeCandidateReview.ps1')
            },
            [pscustomobject]@{
                Module = 'Tools/TerminalEngine/GhosttyRuntimeStatus.psm1'
                Exports = @(
                    'New-RSGhosttyRuntimeStatusReport',
                    'Write-RSGhosttyRuntimeStatusReport')
                Consumers = @('Tools/Get-GhosttyTerminalRuntimeStatus.ps1')
            },
            [pscustomobject]@{
                Module = 'Tools/TerminalEngine/TerminalEngineArchive.psm1'
                Exports = @('Assert-RSTerminalEngineArchiveLayout', 'Get-RSTerminalEngineArchivePlan', 'Expand-RSTerminalEngineArchive')
                Consumers = @('Tools/TerminalEngine/GhosttyRuntimeLifecycle.psm1')
            },
            [pscustomobject]@{
                Module = 'Tools/TerminalEngine/TerminalEngineNotices.psm1'
                Exports = @('Invoke-RSTerminalNoticeMaterialization')
                Consumers = @('Tools/New-TerminalEngineNotices.ps1')
            },
            [pscustomobject]@{
                Module = 'Tools/TerminalEngine/TerminalEnginePeIdentity.psm1'
                Exports = @('Get-RSPeIdentity', 'Get-RSPeSemanticSha256')
                Consumers = @('Tools/TerminalEngine/GhosttyRuntimeLifecycle.psm1')
            }
        )
        $failures = [System.Collections.Generic.List[string]]::new()
        foreach ($contract in $contracts) {
            $modulePath = Join-Path $script:RepoRoot $contract.Module.Replace('/', '\')
            $module = Import-Module $modulePath -Force -PassThru
            foreach ($functionName in $contract.Exports) {
                if (-not $module.ExportedFunctions.ContainsKey($functionName)) {
                    $failures.Add("$($contract.Module): expected current export '$functionName' was not found.")
                }
            }
            foreach ($consumer in $contract.Consumers) {
                $consumerSource = Get-Content -LiteralPath (Join-Path $script:RepoRoot $consumer.Replace('/', '\')) -Raw
                if ($consumerSource -notmatch [regex]::Escape([IO.Path]::GetFileName($contract.Module))) {
                    $failures.Add("${consumer}: no longer references current helper $($contract.Module).")
                }
            }
        }

        Assert-RSNoToolingGovernanceFailures -Failures @($failures) `
            -Contract 'Active Terminal helpers require a current-profile import/export contract.'
    }

    It 'keeps destructive support commands within their documented static safety boundaries' {
        $failures = [System.Collections.Generic.List[string]]::new()

        $iconPath = Join-Path $script:RepoRoot 'Tools\GenerateDebugIcons.ps1'
        $iconSource = Get-Content -LiteralPath $iconPath -Raw
        $iconAst = Get-RSPowerShellAst -Path $iconPath
        $iconStrings = @($iconAst.FindAll({
                    param($node)
                    $node -is [System.Management.Automation.Language.StringConstantExpressionAst] -or
                    $node -is [System.Management.Automation.Language.ExpandableStringExpressionAst]
                }, $true) | ForEach-Object Value)
        $declaredDebugOutputs = @($iconStrings | Where-Object {
                $_ -match '(?i)_debug\.ico$'
            } | Sort-Object -Unique)
        $expectedDebugOutputs = @(
            'RedSalamander\\res\\logo_debug.ico',
            'RedSalamanderMonitor\\res\\logo_background_debug.ico'
        )
        $outputDelta = @(Compare-Object -ReferenceObject $expectedDebugOutputs -DifferenceObject $declaredDebugOutputs)
        if ($outputDelta.Count -ne 0) {
            $failures.Add("Tools/GenerateDebugIcons.ps1: executable source must declare exactly the two reviewed tracked debug-icon outputs.")
        }
        $trackedOutputPaths = @($expectedDebugOutputs | ForEach-Object { $_ -replace '\\+', '/' })
        $trackedDebugOutputs = @(& git -C $script:RepoRoot ls-files -- $trackedOutputPaths)
        if ($LASTEXITCODE -ne 0 -or $trackedDebugOutputs.Count -ne 2) {
            $failures.Add('Tools/GenerateDebugIcons.ps1: both reviewed debug-icon outputs must remain tracked.')
        }
        if ($iconSource -notmatch '(?s)if\s*\(\s*-not\s*\(Test-Path\s+\$InputIconPath\)\s*\)\s*\{\s*throw\s+"Input icon not found:') {
            $failures.Add('Tools/GenerateDebugIcons.ps1: required input icons must be rejected before output generation.')
        }

        $credentialPath = Join-Path $script:RepoRoot 'Tools\Manage-WindowsCredentials.ps1'
        $credentialSource = Get-Content -LiteralPath $credentialPath -Raw
        $credentialAst = Get-RSPowerShellAst -Path $credentialPath
        if ($null -eq $credentialAst.ParamBlock -or
            @($credentialAst.ParamBlock.Attributes | Where-Object {
                    $_.Extent.Text -match '(?i)CmdletBinding\s*\(\s*SupportsShouldProcess'
                }).Count -ne 1) {
            $failures.Add('Tools/Manage-WindowsCredentials.ps1: credential mutation must remain guarded by SupportsShouldProcess.')
        }
        if ($credentialSource -notmatch "(?m)^\s*'\^RedSalamander/'\s*# Safe default:" -or
            $credentialSource -notmatch '\(\$IncludeAllTypes\s+-or\s+\$_\.Type\s+-eq\s+1\)') {
            $failures.Add('Tools/Manage-WindowsCredentials.ps1: default selection must remain RedSalamander-namespaced and Generic-only.')
        }
        if ($credentialSource -notmatch '(?s)if\s*\(\s*-not\s+\$Remove\s*\).*?return.*?\$PSCmdlet\.ShouldProcess\(') {
            $failures.Add('Tools/Manage-WindowsCredentials.ps1: deletion must require both -Remove and ShouldProcess approval.')
        }

        Assert-RSNoToolingGovernanceFailures -Failures @($failures) `
            -Contract 'Support commands with destructive capability require static safety contracts.'
    }

    It 'uses psm1 modules with explicit reviewed exports' {
        $moduleFiles = @(Get-ChildItem -LiteralPath (Join-Path $script:RepoRoot 'Tools\Modules') `
                -Recurse -File)
        $failures = [System.Collections.Generic.List[string]]::new()
        foreach ($file in $moduleFiles) {
            $relativePath = $file.FullName.Substring($script:RepoRoot.Length + 1).Replace('\', '/')
            if ($file.Extension -cne '.psm1') {
                $failures.Add("${relativePath}: files beneath Tools/Modules must be .psm1 modules.")
                continue
            }

            $ast = Get-RSPowerShellAst -Path $file.FullName
            $exports = @($ast.FindAll({
                        param($node)
                        $node -is [System.Management.Automation.Language.CommandAst] -and
                        $node.GetCommandName() -eq 'Export-ModuleMember'
                    }, $true))
            if ($exports.Count -eq 0) {
                $failures.Add("${relativePath}: module has no explicit Export-ModuleMember declaration.")
                continue
            }
            foreach ($export in $exports) {
                if ($export.Extent.Text -notmatch '(?i)(?:^|\s)-Function(?:\s|$)') {
                    $failures.Add("${relativePath}: Export-ModuleMember must explicitly declare -Function exports.")
                }
                if ($export.Extent.Text -match '(?i)-Function\s+[''"]?\*') {
                    $failures.Add("${relativePath}: wildcard module exports are forbidden.")
                }
            }
        }

        Assert-RSNoToolingGovernanceFailures -Failures @($failures) `
            -Contract 'Tools modules require explicit non-wildcard function exports.'
    }

    It 'records every reviewed public naming exception in the inventory' {
        $manifest = Get-RSToolingGovernanceManifest
        $failures = [System.Collections.Generic.List[string]]::new()
        foreach ($entry in @($manifest.entries | Where-Object {
                    $_.visibility -eq 'public' -and [string]$_.path -like '*.ps1'
                })) {
            $commandName = [IO.Path]::GetFileNameWithoutExtension([string]$entry.path)
            $isCanonical = $commandName -cmatch '^[A-Z][A-Za-z0-9]*-[A-Z][A-Za-z0-9-]*$'
            $namingPolicyProperty = $entry.PSObject.Properties['namingPolicy']
            $hasPolicy = $null -ne $namingPolicyProperty -and
                -not [string]::IsNullOrWhiteSpace([string]$namingPolicyProperty.Value)
            if (-not $isCanonical -and -not $hasPolicy) {
                $failures.Add("$($entry.path): noncanonical public command requires reviewed namingPolicy metadata.")
            }
            if ($isCanonical -and $hasPolicy) {
                $failures.Add("$($entry.path): canonical command must not carry naming-exception metadata.")
            }
        }

        Assert-RSNoToolingGovernanceFailures -Failures @($failures) `
            -Contract 'Public naming exceptions must be explicit and narrowly scoped.'
    }
}
