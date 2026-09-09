Set-StrictMode -Version Latest

Import-Module (Join-Path $PSScriptRoot 'TestInvocation.psm1') -Force -Scope Local

function Get-RSTestQuarantineDate {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Value
    )

    if ([string]::IsNullOrWhiteSpace($Value)) {
        return $null
    }

    try {
        return [datetime]::ParseExact($Value, 'yyyy-MM-dd', [System.Globalization.CultureInfo]::InvariantCulture).Date
    } catch {
        return $null
    }
}

function Get-RSTestQuarantineStatus {
    param(
        [object[]]$Entries = @(),

        [datetime]$Now = (Get-Date),

        [int]$MaxDays = 30
    )

    $errors = @()
    $activeEntries = @()
    $invalidEntries = @()
    $requiredFields = @(
        'harness',
        'name',
        'owner',
        'opened',
        'expires',
        'issue',
        'root_cause_hypothesis',
        'fix_or_replace_plan'
    )
    $today = $Now.ToUniversalTime().Date
    $index = 0

    foreach ($entry in @($Entries)) {
        ++$index
        $entryErrors = @()
        foreach ($field in $requiredFields) {
            $value = [string](Get-RSObjectValue -Object $entry -Names @($field) -DefaultValue '')
            if ([string]::IsNullOrWhiteSpace($value)) {
                $entryErrors += "entry $index missing $field"
            }
        }

        $openedText = [string](Get-RSObjectValue -Object $entry -Names @('opened') -DefaultValue '')
        $expiresText = [string](Get-RSObjectValue -Object $entry -Names @('expires') -DefaultValue '')
        $openedDate = Get-RSTestQuarantineDate -Value $openedText
        $expiresDate = Get-RSTestQuarantineDate -Value $expiresText

        if ($null -eq $openedDate) {
            $entryErrors += "entry $index opened must use YYYY-MM-DD"
        }
        if ($null -eq $expiresDate) {
            $entryErrors += "entry $index expires must use YYYY-MM-DD"
        }
        if ($null -ne $openedDate -and $openedDate -gt $today) {
            $entryErrors += "entry $index opened is in the future"
        }
        if ($null -ne $expiresDate -and $expiresDate -lt $today) {
            $entryErrors += "entry $index expired on $($expiresDate.ToString('yyyy-MM-dd'))"
        }
        if ($null -ne $openedDate -and $null -ne $expiresDate) {
            if ($expiresDate -lt $openedDate) {
                $entryErrors += "entry $index expires before opened"
            }
            if (($expiresDate - $openedDate).TotalDays -gt $MaxDays) {
                $entryErrors += "entry $index expires more than $MaxDays days after opened"
            }
        }

        if (@($entryErrors).Count -gt 0) {
            $errors += $entryErrors
            $invalidEntries += $entry
        } else {
            $activeEntries += $entry
        }
    }

    [pscustomobject]@{
        IsValid = (@($errors).Count -eq 0)
        HasBlockingEntries = (@($errors).Count -gt 0 -or @($activeEntries).Count -gt 0)
        TotalCount = @($Entries).Count
        ActiveEntries = @($activeEntries)
        InvalidEntries = @($invalidEntries)
        Errors = @($errors)
    }
}

function Read-RSTestQuarantineFile {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path,

        [datetime]$Now = (Get-Date)
    )

    if (-not (Test-Path -LiteralPath $Path)) {
        return (Get-RSTestQuarantineStatus -Entries @() -Now $Now)
    }

    $entries = @()
    $errors = @()
    $lineNumber = 0
    foreach ($line in @(Get-Content -LiteralPath $Path)) {
        ++$lineNumber
        if ([string]::IsNullOrWhiteSpace($line)) {
            continue
        }

        try {
            $entries += ($line | ConvertFrom-Json)
        } catch {
            $errors += "line $lineNumber is not valid JSON: $($_.Exception.Message)"
        }
    }

    $status = Get-RSTestQuarantineStatus -Entries $entries -Now $Now
    if (@($errors).Count -gt 0) {
        $allErrors = @($status.Errors) + $errors
        $status = [pscustomobject]@{
            IsValid = $false
            HasBlockingEntries = $true
            TotalCount = $status.TotalCount
            ActiveEntries = @($status.ActiveEntries)
            InvalidEntries = @($status.InvalidEntries)
            Errors = @($allErrors)
        }
    }

    return $status
}

function Get-RSTestQuarantineRepairPlan {
    param(
        [AllowNull()]
        [object]$QuarantineStatus = $null,

        [object[]]$TestPlan = @(),

        [hashtable]$HarnessCaseMap = @{}
    )

    $activeEntries = @()
    $invalidEntries = @()
    $errors = @()
    $statusIsValid = $true
    $statusHasBlockingEntries = $false

    if ($null -ne $QuarantineStatus) {
        $activeEntries = @(Get-RSObjectValue -Object $QuarantineStatus -Names @('ActiveEntries', 'active_entries') -DefaultValue @())
        $invalidEntries = @(Get-RSObjectValue -Object $QuarantineStatus -Names @('InvalidEntries', 'invalid_entries') -DefaultValue @())
        $errors = @(Get-RSObjectValue -Object $QuarantineStatus -Names @('Errors', 'errors') -DefaultValue @())
        $statusIsValid = [bool](Get-RSObjectValue -Object $QuarantineStatus -Names @('IsValid', 'is_valid') -DefaultValue $true)
        $statusHasBlockingEntries = [bool](Get-RSObjectValue -Object $QuarantineStatus -Names @('HasBlockingEntries', 'has_blocking_entries') -DefaultValue $false)
    }

    $attempts = @()
    $resolvedInvalidEntries = @($invalidEntries)
    $resolvedErrors = @($errors)

    foreach ($entry in @($activeEntries)) {
        $harness = [string](Get-RSObjectValue -Object $entry -Names @('harness') -DefaultValue '')
        $name = [string](Get-RSObjectValue -Object $entry -Names @('name') -DefaultValue '')
        $matchingPlanEntry = @($TestPlan | Where-Object {
                [string]::Equals(
                    [string](Get-RSObjectValue -Object $_ -Names @('Name', 'name') -DefaultValue ''),
                    $harness,
                    [System.StringComparison]::OrdinalIgnoreCase)
            } | Select-Object -First 1)

        if (@($matchingPlanEntry).Count -eq 0) {
            $resolvedInvalidEntries += $entry
            $resolvedErrors += "quarantine entry [$harness] $name has no matching case in the harness adapter"
            continue
        }

        $planEntry = $matchingPlanEntry[0]
        $kind = [string](Get-RSObjectValue -Object $planEntry -Names @('Kind', 'kind') -DefaultValue '')
        $mode = if ($kind -eq 'SelfTest') { 'selftest-case' } else { 'entry' }
        if ($mode -eq 'selftest-case' -and $null -ne $HarnessCaseMap -and $HarnessCaseMap.ContainsKey($harness)) {
            $knownCaseNames = @($HarnessCaseMap[$harness])
            $caseMatches = @($knownCaseNames | Where-Object {
                    [string]::Equals([string]$_, $name, [System.StringComparison]::OrdinalIgnoreCase)
                })
            if (@($caseMatches).Count -eq 0) {
                $resolvedInvalidEntries += $entry
                $resolvedErrors += "quarantine entry [$harness] $name has no matching case in the harness adapter"
                continue
            }
        }

        $arguments = @($planEntry.Arguments)
        if ($mode -eq 'selftest-case') {
            $arguments = @($arguments | Where-Object { $_ -notlike '--selftest-case=*' })
            $arguments += "--selftest-case=$name"
        }

        $attempts += [pscustomobject]@{
            id = [string](Get-RSObjectValue -Object $planEntry -Names @('Id', 'id') -DefaultValue '')
            harness = $harness
            name = $name
            mode = $mode
            kind = $kind
            path = [string](Get-RSObjectValue -Object $planEntry -Names @('Path', 'path') -DefaultValue '')
            arguments = @($arguments)
            working_directory = [string](Get-RSObjectValue -Object $planEntry -Names @('WorkingDirectory', 'working_directory') -DefaultValue '')
            json_name = [string](Get-RSObjectValue -Object $planEntry -Names @('JsonName', 'json_name') -DefaultValue '')
            requires_interactive_desktop = [bool](Get-RSObjectValue -Object $planEntry -Names @('RequiresInteractiveDesktop', 'requires_interactive_desktop') -DefaultValue $false)
            impact_tags = @((Get-RSObjectValue -Object $planEntry -Names @('ImpactTags', 'impact_tags') -DefaultValue @()))
            input_sets = @((Get-RSObjectValue -Object $planEntry -Names @('InputSets', 'input_sets') -DefaultValue @()))
            artifact_ids = @((Get-RSObjectValue -Object $planEntry -Names @('ArtifactIds', 'artifact_ids') -DefaultValue @()))
            artifact_read_ids = @((Get-RSObjectValue -Object $planEntry -Names @('ArtifactReadIds', 'artifact_read_ids') -DefaultValue @()))
            artifact_write_ids = @((Get-RSObjectValue -Object $planEntry -Names @('ArtifactWriteIds', 'artifact_write_ids') -DefaultValue @()))
            produces_artifact_ids = @((Get-RSObjectValue -Object $planEntry -Names @('ProducesArtifactIds', 'produces_artifact_ids') -DefaultValue @()))
            invalidates_receipt_scopes = @((Get-RSObjectValue -Object $planEntry -Names @('InvalidatesReceiptScopes', 'invalidates_receipt_scopes') -DefaultValue @()))
            cache_policy = [string](Get-RSObjectValue -Object $planEntry -Names @('CachePolicy', 'cache_policy') -DefaultValue '')
            environment_inputs = @((Get-RSObjectValue -Object $planEntry -Names @('EnvironmentInputs', 'environment_inputs') -DefaultValue @()))
            coverage_contract = Get-RSObjectValue -Object $planEntry -Names @('CoverageContract', 'coverage_contract') -DefaultValue $null
            outcome_contract = Get-RSObjectValue -Object $planEntry -Names @('OutcomeContract', 'outcome_contract') -DefaultValue $null
            order_contract = [string](Get-RSObjectValue -Object $planEntry -Names @('OrderContract', 'order_contract') -DefaultValue '')
            resource_classes = @((Get-RSObjectValue -Object $planEntry -Names @('ResourceClasses', 'resource_classes') -DefaultValue @()))
            estimated_duration_key = [string](Get-RSObjectValue -Object $planEntry -Names @('EstimatedDurationKey', 'estimated_duration_key') -DefaultValue '')
            owner = [string](Get-RSObjectValue -Object $entry -Names @('owner') -DefaultValue '')
            opened = [string](Get-RSObjectValue -Object $entry -Names @('opened') -DefaultValue '')
            expires = [string](Get-RSObjectValue -Object $entry -Names @('expires') -DefaultValue '')
            issue = [string](Get-RSObjectValue -Object $entry -Names @('issue') -DefaultValue '')
            root_cause_hypothesis = [string](Get-RSObjectValue -Object $entry -Names @('root_cause_hypothesis') -DefaultValue '')
            fix_or_replace_plan = [string](Get-RSObjectValue -Object $entry -Names @('fix_or_replace_plan') -DefaultValue '')
        }
    }

    [pscustomobject]@{
        IsValid = ($statusIsValid -and @($resolvedInvalidEntries).Count -eq 0 -and @($resolvedErrors).Count -eq 0)
        HasBlockingEntries = ($statusHasBlockingEntries -or @($attempts).Count -gt 0 -or @($resolvedInvalidEntries).Count -gt 0)
        Attempts = @($attempts)
        InvalidEntries = @($resolvedInvalidEntries)
        Errors = @($resolvedErrors)
    }
}

Export-ModuleMember -Function @(
    'Get-RSTestQuarantineDate',
    'Get-RSTestQuarantineStatus',
    'Read-RSTestQuarantineFile',
    'Get-RSTestQuarantineRepairPlan'
)
