Set-StrictMode -Version Latest

$script:RSToolHistoricalManualEntryPoints = [string[]]@(
    'Tools/TerminalEngine/Invoke-HistoricalTerminalTests.ps1',
    'Tools/Evaluate-TerminalEngines.ps1',
    'Tools/Test-TerminalEngineRound4Closeout.ps1'
)

function ConvertTo-RSToolForwardSlashPath {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    return $Path.Replace('\', '/').TrimStart('./')
}

function New-RSToolInventoryFinding {
    param(
        [Parameter(Mandatory = $true)][string]$Code,
        [Parameter(Mandatory = $true)][string]$Path,
        [Parameter(Mandatory = $true)][string]$Message
    )

    return [pscustomobject]@{
        Severity = 'Error'
        Code = $Code
        Path = $Path
        Message = $Message
    }
}

function Get-RSToolInventoryMetadataFindings {
    param(
        [Parameter(Mandatory = $true)]$Item,
        [Parameter(Mandatory = $true)][string]$Identity
    )

    $findings = [System.Collections.Generic.List[object]]::new()
    $requiredScalarProperties = @('kind', 'visibility', 'lifecycle', 'purpose', 'usage', 'ownerSpec')
    $requiredArrayProperties = @('usedBy', 'prerequisites', 'sideEffects', 'outputs', 'tests')

    foreach ($propertyName in $requiredScalarProperties) {
        $property = $Item.PSObject.Properties[$propertyName]
        if ($null -eq $property -or [string]::IsNullOrWhiteSpace([string]$property.Value)) {
            $findings.Add((New-RSToolInventoryFinding `
                        -Code 'INVALID_METADATA' `
                        -Path $Identity `
                        -Message "Required metadata '$propertyName' is missing or empty."))
        }
    }

    foreach ($propertyName in $requiredArrayProperties) {
        $property = $Item.PSObject.Properties[$propertyName]
        if ($null -eq $property) {
            $findings.Add((New-RSToolInventoryFinding `
                        -Code 'INVALID_METADATA' `
                        -Path $Identity `
                        -Message "Required metadata array '$propertyName' is missing."))
        }
    }

    if ($Item.PSObject.Properties['usedBy'] -and @($Item.usedBy).Count -eq 0) {
        $findings.Add((New-RSToolInventoryFinding `
                    -Code 'INVALID_METADATA' `
                    -Path $Identity `
                    -Message "Metadata 'usedBy' must name at least one consumer or usage audience."))
    }

    if ($Item.PSObject.Properties['replacement'] -eq $null) {
        $findings.Add((New-RSToolInventoryFinding `
                    -Code 'INVALID_METADATA' `
                    -Path $Identity `
                    -Message "Required metadata 'replacement' is missing; use null when it does not apply."))
    }

    $lifecycle = if ($Item.PSObject.Properties['lifecycle']) { [string]$Item.lifecycle } else { '' }
    $kind = if ($Item.PSObject.Properties['kind']) { [string]$Item.kind } else { '' }
    if ($lifecycle -eq 'historical') {
        $manualEntryPointsProperty = $Item.PSObject.Properties['manualEntryPoints']
        if ($null -eq $manualEntryPointsProperty -or
            $manualEntryPointsProperty.Value -is [string] -or
            $manualEntryPointsProperty.Value -isnot [System.Collections.IEnumerable]) {
            $findings.Add((New-RSToolInventoryFinding `
                        -Code 'ORPHANED_HISTORICAL_ENTRY' `
                        -Path $Identity `
                        -Message "Historical entries require a non-empty manualEntryPoints array."))
        }
        elseif (@($manualEntryPointsProperty.Value).Count -eq 0) {
            $findings.Add((New-RSToolInventoryFinding `
                        -Code 'ORPHANED_HISTORICAL_ENTRY' `
                        -Path $Identity `
                        -Message "Historical entries require at least one reviewed manual entry point."))
        }
    }

    $requiresSupportPolicy =
        $kind -eq 'compatibility' -or
        $lifecycle -eq 'compatibility'
    $supportPolicyProperty = $Item.PSObject.Properties['supportPolicy']
    if ($requiresSupportPolicy -and
        ($null -eq $supportPolicyProperty -or
            [string]::IsNullOrWhiteSpace([string]$supportPolicyProperty.Value))) {
        $findings.Add((New-RSToolInventoryFinding `
                    -Code 'INVALID_METADATA' `
                    -Path $Identity `
                    -Message "Compatibility records require a non-empty supportPolicy."))
    }
    elseif ($null -ne $supportPolicyProperty -and
        $null -ne $supportPolicyProperty.Value -and
        [string]::IsNullOrWhiteSpace([string]$supportPolicyProperty.Value)) {
        $findings.Add((New-RSToolInventoryFinding `
                    -Code 'INVALID_METADATA' `
                    -Path $Identity `
                    -Message "Optional metadata 'supportPolicy' must be non-empty when present."))
    }

    $namingPolicyProperty = $Item.PSObject.Properties['namingPolicy']
    if ($null -ne $namingPolicyProperty -and
        $null -ne $namingPolicyProperty.Value -and
        [string]::IsNullOrWhiteSpace([string]$namingPolicyProperty.Value)) {
        $findings.Add((New-RSToolInventoryFinding `
                    -Code 'INVALID_METADATA' `
                    -Path $Identity `
                    -Message "Optional metadata 'namingPolicy' must be non-empty when present."))
    }

    $itemPath = if ($Item.PSObject.Properties['path']) { [string]$Item.path } else { '' }
    $isPublicPowerShellCommand =
        [string]$Item.visibility -eq 'public' -and
        [string]$Item.kind -in @('command', 'compatibility') -and
        $itemPath.EndsWith('.ps1', [StringComparison]::OrdinalIgnoreCase)
    if ($isPublicPowerShellCommand) {
        $leafName = [IO.Path]::GetFileName($itemPath)
        if ($leafName -notmatch '^[^-]+-[^-]+\.ps1$' -and
            ($null -eq $namingPolicyProperty -or
                [string]::IsNullOrWhiteSpace([string]$namingPolicyProperty.Value))) {
            $findings.Add((New-RSToolInventoryFinding `
                        -Code 'INVALID_METADATA' `
                        -Path $Identity `
                        -Message "Public non-Verb-Noun entrypoint '$leafName' requires a reviewed namingPolicy."))
        }
    }

    if ($Item.PSObject.Properties['kind'] -and
        [string]$Item.kind -notin @('command', 'module', 'test', 'data', 'generated', 'compatibility', 'documentation', 'source', 'project', 'patch')) {
        $findings.Add((New-RSToolInventoryFinding `
                    -Code 'INVALID_METADATA' `
                    -Path $Identity `
                    -Message "Unsupported kind '$($Item.kind)'."))
    }
    if ($Item.PSObject.Properties['visibility'] -and [string]$Item.visibility -notin @('public', 'internal')) {
        $findings.Add((New-RSToolInventoryFinding `
                    -Code 'INVALID_METADATA' `
                    -Path $Identity `
                    -Message "Unsupported visibility '$($Item.visibility)'."))
    }
    if ($Item.PSObject.Properties['lifecycle'] -and
        [string]$Item.lifecycle -notin @('active', 'compatibility', 'historical', 'generated', 'data')) {
        $findings.Add((New-RSToolInventoryFinding `
                    -Code 'INVALID_METADATA' `
                    -Path $Identity `
                    -Message "Unsupported lifecycle '$($Item.lifecycle)'."))
    }

    $consumerClosureProperty = $Item.PSObject.Properties['consumerClosure']
    if ($null -ne $consumerClosureProperty -and
        [string]$consumerClosureProperty.Value -ne 'derived-references') {
        $findings.Add((New-RSToolInventoryFinding `
                    -Code 'INVALID_METADATA' `
                    -Path $Identity `
                    -Message "Unsupported consumerClosure '$($consumerClosureProperty.Value)'."))
    }

    return @($findings)
}

function Get-RSToolRepositoryFilePaths {
    param(
        [Parameter(Mandatory = $true)]
        [string]$RepositoryRoot
    )

    $gitOutput = @(& git -C $RepositoryRoot ls-files --cached --others --exclude-standard -- Tools 2>&1)
    if ($LASTEXITCODE -ne 0) {
        throw "Failed to enumerate tracked and non-ignored Tools files from '$RepositoryRoot': $($gitOutput -join '; ')"
    }

    return @($gitOutput |
        ForEach-Object { ConvertTo-RSToolForwardSlashPath -Path ([string]$_) } |
        Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
        Where-Object {
            $fullPath = Join-Path $RepositoryRoot $_
            Test-Path -LiteralPath $fullPath -PathType Leaf
        } |
        Sort-Object -Unique)
}

function Get-RSToolDerivedConsumers {
    param(
        [Parameter(Mandatory = $true)][string]$RepositoryRoot,
        [Parameter(Mandatory = $true)][string]$ToolPath
    )

    $leaf = [IO.Path]::GetFileName($ToolPath)
    $candidatePaths = @(& git -C $RepositoryRoot ls-files --cached --others --exclude-standard)
    if ($LASTEXITCODE -ne 0) { throw 'Failed to enumerate repository consumer inputs.' }
    $consumers = foreach ($candidatePath in $candidatePaths) {
        $relative = ConvertTo-RSToolForwardSlashPath -Path ([string]$candidatePath)
        if ($relative -ieq $ToolPath -or $relative -notmatch '\.(?:ps1|psm1|psd1|yml|yaml)$') { continue }
        $fullPath = Join-Path $RepositoryRoot $relative
        if (-not (Test-Path -LiteralPath $fullPath -PathType Leaf)) { continue }
        $source = Get-Content -LiteralPath $fullPath -Raw
        if ($source.IndexOf($leaf, [StringComparison]::OrdinalIgnoreCase) -ge 0) { $relative }
    }

    return @($consumers | Sort-Object -Unique)
}

function Test-RSToolConsumerIsDeclared {
    param(
        [Parameter(Mandatory = $true)][string]$Consumer,
        [Parameter(Mandatory = $true)][object[]]$Declaration
    )

    foreach ($declared in $Declaration) {
        $pattern = [string]$declared
        if ($pattern.IndexOfAny([char[]]'*?[') -ge 0) {
            if ([WildcardPattern]::new($pattern, [WildcardOptions]::IgnoreCase).IsMatch($Consumer)) { return $true }
        } elseif ([string]::Equals($pattern, $Consumer, [StringComparison]::OrdinalIgnoreCase)) {
            return $true
        }
    }
    return $false
}

function Get-RSToolInventory {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$RepositoryRoot,

        [string]$ManifestPath = 'Tools/tool-inventory.json',

        [AllowEmptyCollection()]
        [string[]]$FilePaths
    )

    $repositoryFull = [IO.Path]::GetFullPath($RepositoryRoot)
    $manifestFull = if ([IO.Path]::IsPathRooted($ManifestPath)) {
        [IO.Path]::GetFullPath($ManifestPath)
    }
    else {
        [IO.Path]::GetFullPath((Join-Path $repositoryFull $ManifestPath))
    }
    if (-not (Test-Path -LiteralPath $manifestFull -PathType Leaf)) {
        throw "Tool inventory manifest was not found: $manifestFull"
    }

    try {
        $manifest = Get-Content -LiteralPath $manifestFull -Raw | ConvertFrom-Json
    }
    catch {
        throw "Tool inventory manifest is not valid JSON: $manifestFull. $($_.Exception.Message)"
    }

    $findings = [System.Collections.Generic.List[object]]::new()
    if ($manifest.schemaVersion -ne 3) {
        $findings.Add((New-RSToolInventoryFinding `
                    -Code 'UNSUPPORTED_SCHEMA' `
                    -Path (ConvertTo-RSToolForwardSlashPath -Path $ManifestPath) `
                    -Message "Expected schemaVersion 3 but found '$($manifest.schemaVersion)'."))
    }

    $entries = @($manifest.entries)
    $rules = @($manifest.rules)
    $entryByPath = @{}
    foreach ($entry in $entries) {
        $path = ConvertTo-RSToolForwardSlashPath -Path ([string]$entry.path)
        $identity = if ([string]::IsNullOrWhiteSpace($path)) { '<entry-without-path>' } else { $path }
        foreach ($finding in @(Get-RSToolInventoryMetadataFindings -Item $entry -Identity $identity)) {
            $findings.Add($finding)
        }
        if ([string]::IsNullOrWhiteSpace($path) -or -not $path.StartsWith('Tools/', [StringComparison]::OrdinalIgnoreCase)) {
            $findings.Add((New-RSToolInventoryFinding `
                        -Code 'INVALID_ENTRY_PATH' `
                        -Path $identity `
                        -Message 'Explicit entry paths must be repository-relative paths beneath Tools/.' ))
            continue
        }
        $key = $path.ToUpperInvariant()
        if ($entryByPath.ContainsKey($key)) {
            $findings.Add((New-RSToolInventoryFinding `
                        -Code 'DUPLICATE_ENTRY' `
                        -Path $path `
                        -Message 'The manifest contains more than one explicit entry for this path.'))
            continue
        }
        $entryByPath[$key] = $entry
    }

    $ruleIds = @{}
    foreach ($rule in $rules) {
        $identity = if ([string]::IsNullOrWhiteSpace([string]$rule.id)) { '<rule-without-id>' } else { "rule:$($rule.id)" }
        foreach ($finding in @(Get-RSToolInventoryMetadataFindings -Item $rule -Identity $identity)) {
            $findings.Add($finding)
        }
        if ([string]::IsNullOrWhiteSpace([string]$rule.id) -or [string]::IsNullOrWhiteSpace([string]$rule.pathPattern)) {
            $findings.Add((New-RSToolInventoryFinding `
                        -Code 'INVALID_RULE' `
                        -Path $identity `
                        -Message 'Rules require non-empty id and pathPattern properties.'))
            continue
        }
        $ruleKey = ([string]$rule.id).ToUpperInvariant()
        if ($ruleIds.ContainsKey($ruleKey)) {
            $findings.Add((New-RSToolInventoryFinding `
                        -Code 'DUPLICATE_RULE' `
                        -Path $identity `
                        -Message 'Rule identifiers must be unique.'))
        }
        else {
            $ruleIds[$ruleKey] = $true
        }
        try {
            $null = [regex]::new([string]$rule.pathPattern, [Text.RegularExpressions.RegexOptions]::IgnoreCase)
        }
        catch {
            $findings.Add((New-RSToolInventoryFinding `
                        -Code 'INVALID_RULE' `
                        -Path $identity `
                        -Message "pathPattern is not a valid regular expression: $($_.Exception.Message)"))
        }
    }

    $discoveredPaths = @(if ($PSBoundParameters.ContainsKey('FilePaths')) {
            $FilePaths | ForEach-Object { ConvertTo-RSToolForwardSlashPath -Path $_ } | Sort-Object -Unique
        }
        else {
            Get-RSToolRepositoryFilePaths -RepositoryRoot $repositoryFull
        })
    $discoveredSet = @{}
    foreach ($path in $discoveredPaths) {
        $discoveredSet[$path.ToUpperInvariant()] = $true
    }

    foreach ($entryPath in $entryByPath.Keys) {
        if (-not $discoveredSet.ContainsKey($entryPath)) {
            $entry = $entryByPath[$entryPath]
            $findings.Add((New-RSToolInventoryFinding `
                        -Code 'STALE_ENTRY' `
                        -Path (ConvertTo-RSToolForwardSlashPath -Path ([string]$entry.path)) `
                        -Message 'The manifest entry names a path that is not tracked or present as a non-ignored working-tree file.'))
        }
    }

    foreach ($entry in $entries) {
        if ($null -eq $entry.replacement) {
            continue
        }

        $replacement = ConvertTo-RSToolForwardSlashPath -Path ([string]$entry.replacement)
        if ([string]::IsNullOrWhiteSpace($replacement) -or
            -not $replacement.StartsWith('Tools/', [StringComparison]::OrdinalIgnoreCase) -or
            -not $discoveredSet.ContainsKey($replacement.ToUpperInvariant())) {
            $findings.Add((New-RSToolInventoryFinding `
                        -Code 'INVALID_REPLACEMENT' `
                        -Path (ConvertTo-RSToolForwardSlashPath -Path ([string]$entry.path)) `
                        -Message "Replacement '$($entry.replacement)' must be an exact existing repository-relative Tools path."))
        }
    }

    foreach ($entry in @($entries | Where-Object { [string]$_.lifecycle -eq 'historical' })) {
        $entryPath = ConvertTo-RSToolForwardSlashPath -Path ([string]$entry.path)
        $manualEntryPointsProperty = $entry.PSObject.Properties['manualEntryPoints']
        if ($null -eq $manualEntryPointsProperty -or
            $manualEntryPointsProperty.Value -is [string] -or
            $manualEntryPointsProperty.Value -isnot [System.Collections.IEnumerable]) {
            continue
        }

        $seenManualEntryPoints = @{}
        foreach ($manualEntryPointValue in @($manualEntryPointsProperty.Value)) {
            $manualEntryPointText = [string]$manualEntryPointValue
            $manualEntryPoint = ConvertTo-RSToolForwardSlashPath -Path $manualEntryPointText
            if ([string]::IsNullOrWhiteSpace($manualEntryPointText) -or
                $manualEntryPointText -cne $manualEntryPoint -or
                $script:RSToolHistoricalManualEntryPoints -cnotcontains $manualEntryPoint) {
                $findings.Add((New-RSToolInventoryFinding `
                            -Code 'INVALID_MANUAL_ENTRY_POINT' `
                            -Path $entryPath `
                            -Message "Manual entry point '$manualEntryPointText' is not one of the exact reviewed historical command paths."))
                continue
            }

            $manualEntryPointKey = $manualEntryPoint.ToUpperInvariant()
            if ($seenManualEntryPoints.ContainsKey($manualEntryPointKey)) {
                $findings.Add((New-RSToolInventoryFinding `
                            -Code 'INVALID_MANUAL_ENTRY_POINT' `
                            -Path $entryPath `
                            -Message "Manual entry point '$manualEntryPointText' is declared more than once."))
                continue
            }
            $seenManualEntryPoints[$manualEntryPointKey] = $true

            if (-not $discoveredSet.ContainsKey($manualEntryPointKey)) {
                $findings.Add((New-RSToolInventoryFinding `
                            -Code 'MISSING_MANUAL_ENTRY_POINT' `
                            -Path $entryPath `
                            -Message "Manual entry point '$manualEntryPoint' is not an existing Tools file."))
                continue
            }
            if (-not $entryByPath.ContainsKey($manualEntryPointKey)) {
                $findings.Add((New-RSToolInventoryFinding `
                            -Code 'INVALID_MANUAL_ENTRY_POINT' `
                            -Path $entryPath `
                            -Message "Manual entry point '$manualEntryPoint' has no explicit inventory entry."))
                continue
            }

            $manualEntryPointEntry = $entryByPath[$manualEntryPointKey]
            if ([string]$manualEntryPointEntry.path -cne $manualEntryPoint -or
                [string]$manualEntryPointEntry.lifecycle -cne 'historical' -or
                [string]$manualEntryPointEntry.kind -cne 'command') {
                $findings.Add((New-RSToolInventoryFinding `
                            -Code 'NONHISTORICAL_MANUAL_ENTRY_POINT' `
                            -Path $entryPath `
                            -Message "Manual entry point '$manualEntryPoint' must resolve exactly to a historical command entry."))
            }
        }
    }

    foreach ($entry in @($entries | Where-Object {
                $null -ne $_.PSObject.Properties['consumerClosure'] -and
                [string]$_.consumerClosure -eq 'derived-references'
            })) {
        $entryPath = ConvertTo-RSToolForwardSlashPath -Path ([string]$entry.path)
        foreach ($consumer in @(Get-RSToolDerivedConsumers -RepositoryRoot $repositoryFull -ToolPath $entryPath)) {
            if (-not (Test-RSToolConsumerIsDeclared -Consumer $consumer -Declaration @($entry.usedBy))) {
                $findings.Add((New-RSToolInventoryFinding `
                            -Code 'MISSING_DERIVED_CONSUMER' `
                            -Path $entryPath `
                            -Message "Derived repository consumer '$consumer' is missing from usedBy."))
            }
        }
    }

    $resolvedFiles = [System.Collections.Generic.List[object]]::new()
    foreach ($path in $discoveredPaths) {
        $key = $path.ToUpperInvariant()
        $relativeBelowTools = $path.Substring('Tools/'.Length)
        $isTopLevel = -not $relativeBelowTools.Contains('/')
        $metadata = $null
        $classificationSource = $null

        if ($entryByPath.ContainsKey($key)) {
            $metadata = $entryByPath[$key]
            $classificationSource = 'entry'
        }
        else {
            $matchingRules = @($rules | Where-Object {
                    try {
                        [regex]::IsMatch($path, [string]$_.pathPattern, [Text.RegularExpressions.RegexOptions]::IgnoreCase)
                    }
                    catch {
                        $false
                    }
                })
            if ($isTopLevel) {
                $findings.Add((New-RSToolInventoryFinding `
                            -Code 'TOP_LEVEL_REQUIRES_EXPLICIT_ENTRY' `
                            -Path $path `
                            -Message 'Every file directly beneath Tools/ must have an explicit manifest entry; directory rules cannot classify the public root surface.'))
            }
            if ($matchingRules.Count -eq 0) {
                $findings.Add((New-RSToolInventoryFinding `
                            -Code 'UNCLASSIFIED_PATH' `
                            -Path $path `
                            -Message 'No explicit manifest entry or reviewed directory rule classifies this file.'))
                continue
            }
            if ($matchingRules.Count -gt 1) {
                $findings.Add((New-RSToolInventoryFinding `
                            -Code 'AMBIGUOUS_CLASSIFICATION' `
                            -Path $path `
                            -Message "More than one directory rule matches this file: $(@($matchingRules.id) -join ', ')."))
                continue
            }
            $metadata = $matchingRules[0]
            $classificationSource = "rule:$($matchingRules[0].id)"
        }

        $resolvedFiles.Add([pscustomobject]@{
                Path = $path
                Kind = [string]$metadata.kind
                Visibility = [string]$metadata.visibility
                Lifecycle = [string]$metadata.lifecycle
                Purpose = [string]$metadata.purpose
                Usage = [string]$metadata.usage
                UsedBy = @($metadata.usedBy)
                ManualEntryPoints = @(
                    if ($null -ne $metadata.PSObject.Properties['manualEntryPoints']) {
                        @($metadata.manualEntryPoints)
                    }
                )
                Prerequisites = if ($null -eq $metadata.PSObject.Properties['prerequisites']) {
                    @()
                }
                else {
                    @($metadata.prerequisites)
                }
                SideEffects = @($metadata.sideEffects)
                Outputs = @($metadata.outputs)
                OwnerSpec = [string]$metadata.ownerSpec
                Tests = @($metadata.tests)
                Replacement = if ($null -eq $metadata.replacement) { $null } else { [string]$metadata.replacement }
                SupportPolicy = if ($null -eq $metadata.PSObject.Properties['supportPolicy'] -or
                    $null -eq $metadata.supportPolicy) {
                    $null
                }
                else {
                    [string]$metadata.supportPolicy
                }
                NamingPolicy = if ($null -eq $metadata.PSObject.Properties['namingPolicy'] -or
                    $null -eq $metadata.namingPolicy) {
                    $null
                }
                else {
                    [string]$metadata.namingPolicy
                }
                ClassificationSource = $classificationSource
            })
    }

    $sortedFindings = @($findings | Sort-Object Code, Path, Message)
    $sortedFiles = @($resolvedFiles | Sort-Object Path)
    return [pscustomobject]@{
        SchemaVersion = $manifest.schemaVersion
        ManifestPath = $manifestFull
        GeneratedAtUtc = [DateTime]::UtcNow.ToString('o')
        Files = $sortedFiles
        Findings = $sortedFindings
        Summary = [pscustomobject]@{
            DiscoveredFiles = $discoveredPaths.Count
            ClassifiedFiles = $sortedFiles.Count
            ExplicitEntries = $entries.Count
            DirectoryRules = $rules.Count
            BlockingFindings = $sortedFindings.Count
        }
    }
}

function ConvertTo-RSToolInventoryMarkdownCell {
    param([AllowNull()][object]$Value)

    if ($null -eq $Value) {
        return ''
    }
    return ([string]$Value).Replace('|', '\|').Replace("`r", ' ').Replace("`n", ' ')
}

function ConvertTo-RSToolInventoryMarkdown {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        $Inventory
    )

    $lines = [System.Collections.Generic.List[string]]::new()
    $lines.Add('# Tools inventory')
    $lines.Add('')
    $lines.Add("- Discovered files: $($Inventory.Summary.DiscoveredFiles)")
    $lines.Add("- Classified files: $($Inventory.Summary.ClassifiedFiles)")
    $lines.Add("- Blocking findings: $($Inventory.Summary.BlockingFindings)")
    $lines.Add('')
    $lines.Add('## Findings')
    $lines.Add('')
    if (@($Inventory.Findings).Count -eq 0) {
        $lines.Add('None.')
    }
    else {
        $lines.Add('| Code | Path | Message |')
        $lines.Add('|---|---|---|')
        foreach ($finding in $Inventory.Findings) {
            $lines.Add("| $($finding.Code) | ``$(ConvertTo-RSToolInventoryMarkdownCell $finding.Path)`` | $(ConvertTo-RSToolInventoryMarkdownCell $finding.Message) |")
        }
    }
    $lines.Add('')
    $lines.Add('## Files')
    $lines.Add('')
    $lines.Add('| Path | Kind | Visibility | Lifecycle | Purpose | Primary consumers |')
    $lines.Add('|---|---|---|---|---|---|')
    foreach ($file in $Inventory.Files) {
        $consumers = @($file.UsedBy) -join '; '
        $lines.Add("| ``$(ConvertTo-RSToolInventoryMarkdownCell $file.Path)`` | $($file.Kind) | $($file.Visibility) | $($file.Lifecycle) | $(ConvertTo-RSToolInventoryMarkdownCell $file.Purpose) | $(ConvertTo-RSToolInventoryMarkdownCell $consumers) |")
    }

    return $lines -join [Environment]::NewLine
}

Export-ModuleMember -Function Get-RSToolInventory, ConvertTo-RSToolInventoryMarkdown
