Set-StrictMode -Version Latest

function Get-RSSelfTestArguments {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Flag,

        [double]$TimeoutMultiplier = 1.0,

        [switch]$FailFast,

        [string]$CaseFilter = '',

        [string]$CommandsFamily = '',

        [uint32]$RepeatCount = 1,

        [string]$ShuffleSeed = '',

        [string]$PerfBudgetPath = '',

        [switch]$RequirePerfBudgets,

        [switch]$NoActivate,

        [string]$SelfTestFlakyProofCase = '',

        [string]$SelfTestOrderProofCase = ''
    )

    $arguments = @($Flag)
    if ($NoActivate) {
        $arguments += '--selftest-no-activate'
    }
    if ($FailFast) {
        $arguments += '--selftest-fail-fast'
    }
    if ($TimeoutMultiplier -ne 1.0) {
        $timeoutText = [string]::Format([System.Globalization.CultureInfo]::InvariantCulture, '{0:g}', $TimeoutMultiplier)
        $arguments += "--selftest-timeout-multiplier=$timeoutText"
    }
    if (-not [string]::IsNullOrWhiteSpace($CaseFilter)) {
        $arguments += "--selftest-case=$CaseFilter"
    }
    if (-not [string]::IsNullOrWhiteSpace($CommandsFamily)) {
        $arguments += "--selftest-family=$CommandsFamily"
    }
    if ($RepeatCount -gt 1) {
        $arguments += "--selftest-repeat=$RepeatCount"
    }
    if (-not [string]::IsNullOrWhiteSpace($ShuffleSeed)) {
        $arguments += "--selftest-shuffle=$ShuffleSeed"
    }
    if (-not [string]::IsNullOrWhiteSpace($PerfBudgetPath)) {
        $arguments += "--selftest-perf-budget=$PerfBudgetPath"
    }
    if ($RequirePerfBudgets) {
        $arguments += '--selftest-require-perf-budgets'
    }
    if (-not [string]::IsNullOrWhiteSpace($SelfTestFlakyProofCase)) {
        $arguments += "--selftest-flaky-proof-case=$SelfTestFlakyProofCase"
    }
    if (-not [string]::IsNullOrWhiteSpace($SelfTestOrderProofCase)) {
        $arguments += "--selftest-order-proof-case=$SelfTestOrderProofCase"
    }

    return $arguments
}

function Get-RSSelfTestListArguments {
    param(
        [string[]]$SelfTestArguments = @()
    )

    $arguments = @('--selftest-list-cases')
    foreach ($argument in @($SelfTestArguments)) {
        if ($argument -in @('--compare-selftest', '--commands-selftest', '--fileops-selftest') -or
            $argument -like '--selftest-case=*' -or $argument -like '--selftest-family=*') {
            $arguments += $argument
        }
    }

    return $arguments
}

function New-RSPesterInvokeParameters {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path,

        [string[]]$Arguments = @(),

        [object]$InvokePesterCommand = (Get-Command Invoke-Pester -ErrorAction Stop)
    )

    $parameters = @{
        PassThru = $true
    }

    if ($InvokePesterCommand.Parameters.ContainsKey('Path')) {
        $parameters['Path'] = $Path
    } elseif ($InvokePesterCommand.Parameters.ContainsKey('Script')) {
        $parameters['Script'] = $Path
    } else {
        throw 'Invoke-Pester exposes neither -Path nor -Script; cannot run tooling Pester tests.'
    }

    for ($i = 0; $i -lt @($Arguments).Count; $i++) {
        $argument = [string]$Arguments[$i]
        switch ($argument) {
            '-ExcludeTag' {
                if ($i + 1 -ge @($Arguments).Count) {
                    throw 'Pester -ExcludeTag requires a value.'
                }
                $i++
                $existing = @()
                if ($parameters.ContainsKey('ExcludeTag')) {
                    $existing = @($parameters['ExcludeTag'])
                }
                $parameters['ExcludeTag'] = $existing + [string]$Arguments[$i]
            }
            '-Tag' {
                if ($i + 1 -ge @($Arguments).Count) {
                    throw 'Pester -Tag requires a value.'
                }
                $i++
                $existing = @()
                if ($parameters.ContainsKey('Tag')) {
                    $existing = @($parameters['Tag'])
                }
                $parameters['Tag'] = $existing + [string]$Arguments[$i]
            }
            default {
                throw "Unsupported Pester runner argument '$argument'. Add explicit binding support before forwarding it."
            }
        }
    }

    return $parameters
}

function Get-RSBuildScriptArguments {
    param(
        [Parameter(Mandatory = $true)]
        [ValidateSet('All', 'Compare', 'Commands', 'FileOps', 'CI', 'Full')]
        [string]$Suite,

        [Parameter(Mandatory = $true)]
        [string]$Configuration,

        [Parameter(Mandatory = $true)]
        [ValidateSet('x64', 'ARM64')]
        [string]$Platform
    )

    $arguments = @{
        Configuration = $Configuration
        Platform = $Platform
    }

    if ($Suite -notin @('CI', 'Full')) {
        $arguments.ProjectName = 'RedSalamander'
    }
    elseif ($Suite -eq 'Full') {
        # The complete test-enabled solution launches many compiler, linker,
        # resource-compiler, and vcpkg applocal processes. Bound MSBuild fan-out
        # so process initialization remains reliable on interactive hosts.
        $arguments.MaxCpuCount = 4
    }

    return $arguments
}

function Get-RSBuildEnvironmentOverrides {
    param(
        [Parameter(Mandatory = $true)]
        [ValidateSet('All', 'Compare', 'Commands', 'FileOps', 'CI', 'Full')]
        [string]$Suite
    )

    if ($Suite -eq 'Full') {
        return @{
            RSBuildEnableTests = 'true'
            RSBuildCompilerDebugInformation = 'Embedded'
        }
    }

    return @{}
}

function Test-RSSelfTestResultCoverage {
    param(
        [string[]]$ExpectedCaseNames = @(),

        [object[]]$ActualCases = @(),

        [string[]]$AllowedExtraCaseNames = @(),

        [uint32]$ExpectedRepeatCount = 1,

        [switch]$RequireExpectedCases
    )

    $actualNames = @($ActualCases | ForEach-Object {
            $property = $_.PSObject.Properties['name']
            if ($null -ne $property) {
                [string]$property.Value
            }
        } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })

    $expectedGroups = @($ExpectedCaseNames | Group-Object | Where-Object { $_.Count -gt 1 })
    $expectedRepeat = [Math]::Max(1, [int]$ExpectedRepeatCount)

    $expectedSet = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::Ordinal)
    foreach ($name in @($ExpectedCaseNames)) {
        [void]$expectedSet.Add($name)
    }

    $actualGroups = @($actualNames | Group-Object | Where-Object {
            if ($expectedRepeat -le 1) {
                return $_.Count -gt 1
            }

            return ($expectedSet.Contains($_.Name) -and $_.Count -ne $expectedRepeat) -or
                ((-not $expectedSet.Contains($_.Name)) -and $_.Count -gt 1)
        })

    $actualSet = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::Ordinal)
    foreach ($name in @($actualNames)) {
        [void]$actualSet.Add($name)
    }

    $allowedExtras = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::Ordinal)
    foreach ($name in @($AllowedExtraCaseNames)) {
        [void]$allowedExtras.Add($name)
    }

    $missing = @($ExpectedCaseNames | Where-Object { -not $actualSet.Contains($_) } | Sort-Object -Unique)
    $extra = @($actualNames | Where-Object { (-not $expectedSet.Contains($_)) -and (-not $allowedExtras.Contains($_)) } | Sort-Object -Unique)
    $noExpectedCases = ($RequireExpectedCases -and @($ExpectedCaseNames).Count -eq 0)

    [pscustomobject]@{
        IsValid = ((-not $noExpectedCases) -and @($expectedGroups).Count -eq 0 -and @($actualGroups).Count -eq 0 -and @($missing).Count -eq 0 -and @($extra).Count -eq 0)
        NoExpectedCases = [bool]$noExpectedCases
        DuplicateExpected = @($expectedGroups | Select-Object -ExpandProperty Name)
        DuplicateActual = @($actualGroups | Select-Object -ExpandProperty Name)
        Missing = @($missing)
        Extra = @($extra)
    }
}

function Get-RSObjectValue {
    param(
        [Parameter(Mandatory = $true)]
        [AllowNull()]
        [object]$Object,

        [Parameter(Mandatory = $true)]
        [string[]]$Names,

        [AllowNull()]
        [object]$DefaultValue = $null
    )

    if ($null -eq $Object) {
        return $DefaultValue
    }

    if ($Object -is [System.Collections.IDictionary]) {
        foreach ($name in $Names) {
            if ($Object.Contains($name)) {
                return $Object[$name]
            }
        }
        return $DefaultValue
    }

    foreach ($name in $Names) {
        $property = $Object.PSObject.Properties[$name]
        if ($null -ne $property) {
            return $property.Value
        }
    }

    return $DefaultValue
}

function Set-RSObjectValue {
    param(
        [Parameter(Mandatory = $true)]
        [object]$Object,

        [Parameter(Mandatory = $true)]
        [string]$Name,

        [AllowNull()]
        [object]$Value = $null
    )

    if ($Object -is [System.Collections.IDictionary]) {
        $Object[$Name] = $Value
        return $Object
    }

    $property = $Object.PSObject.Properties[$Name]
    if ($null -ne $property) {
        $property.Value = $Value
    } else {
        $Object | Add-Member -NotePropertyName $Name -NotePropertyValue $Value -Force
    }

    return $Object
}

function Test-RSTestResultPassed {
    param(
        [Parameter(Mandatory = $true)]
        [object]$Result
    )

    $exitCode = [int](Get-RSObjectValue -Object $Result -Names @('ExitCode', 'exit_code') -DefaultValue 1)
    $failed = [int](Get-RSObjectValue -Object $Result -Names @('Failed', 'failed') -DefaultValue 0)
    return ($exitCode -eq 0 -and $failed -eq 0)
}

function Add-RSTestResultClassification {
    param(
        [Parameter(Mandatory = $true)]
        [object]$Result,

        [object[]]$RetryResults = @()
    )

    $retryAttempts = @($RetryResults)
    $classification = ''
    $classificationReason = ''

    if (Test-RSTestResultPassed -Result $Result) {
        $classification = 'PASSED'
        $classificationReason = 'Initial attempt passed.'
    } elseif (@($retryAttempts).Count -eq 0) {
        $classification = 'UNCLASSIFIED_FAILURE'
        $classificationReason = 'Initial attempt failed and no retry classification evidence was collected.'
    } else {
        $shuffleTriageAttempts = @($retryAttempts | Where-Object {
                [string](Get-RSObjectValue -Object $_ -Names @('mode', 'Mode') -DefaultValue '') -eq 'shuffle-triage'
            })
        $hasShuffleTriage = @($shuffleTriageAttempts).Count -gt 0
        $hasFailedShuffleTriage = @($shuffleTriageAttempts | Where-Object {
                -not [bool](Get-RSObjectValue -Object $_ -Names @('passed', 'Passed') -DefaultValue $false)
            }).Count -gt 0

        $allRetriesPassed = $true
        foreach ($retryAttempt in $retryAttempts) {
            $retryPassed = [bool](Get-RSObjectValue -Object $retryAttempt -Names @('passed', 'Passed') -DefaultValue $false)
            if (-not $retryPassed) {
                $allRetriesPassed = $false
                break
            }
        }

        if ($hasFailedShuffleTriage) {
            $classification = 'REGRESSION'
            $classificationReason = 'Initial broad self-test attempt failed, failed cases passed when rerun in isolation, but shuffle triage reproduced a failure; this is a blocking isolation/order regression.'
        } elseif ($allRetriesPassed) {
            $kind = [string](Get-RSObjectValue -Object $Result -Names @('Kind', 'kind') -DefaultValue '')
            $nonShuffleAttempts = @($retryAttempts | Where-Object {
                    [string](Get-RSObjectValue -Object $_ -Names @('mode', 'Mode') -DefaultValue '') -ne 'shuffle-triage'
                })
            $allFailedCaseRetries = @($nonShuffleAttempts).Count -gt 0
            foreach ($retryAttempt in $nonShuffleAttempts) {
                $mode = [string](Get-RSObjectValue -Object $retryAttempt -Names @('mode', 'Mode') -DefaultValue '')
                if ($mode -ne 'failed-case') {
                    $allFailedCaseRetries = $false
                    break
                }
            }

            if ($kind -eq 'SelfTest' -and $allFailedCaseRetries -and $hasShuffleTriage) {
                $classification = 'FLAKY'
                $classificationReason = 'Initial broad self-test attempt failed, failed cases passed when rerun in isolation, and shuffle triage passed; this remains a blocking flaky-test repair item.'
            } elseif ($kind -eq 'SelfTest' -and $allFailedCaseRetries) {
                $classification = 'ISOLATION_SUSPECT'
                $classificationReason = 'Initial broad self-test attempt failed, but failed cases passed when rerun in isolation; shuffle triage is still required.'
            } else {
                $classification = 'FLAKY'
                $classificationReason = 'Initial attempt failed, but retry evidence passed; this remains a blocking flaky-test repair item.'
            }
        } else {
            $classification = 'REGRESSION'
            $classificationReason = 'Initial attempt failed and retry evidence failed again.'
        }
    }

    $Result = Set-RSObjectValue -Object $Result -Name 'Classification' -Value $classification
    $Result = Set-RSObjectValue -Object $Result -Name 'ClassificationReason' -Value $classificationReason
    $Result = Set-RSObjectValue -Object $Result -Name 'RetryAttempts' -Value $retryAttempts

    return $Result
}

function Get-RSSelfTestClassificationRetryPlan {
    param(
        [Parameter(Mandatory = $true)]
        [object]$Entry,

        [string[]]$FailureNames = @(),

        [string[]]$ShuffleTriageSeeds = @()
    )

    $entryArguments = @($Entry.Arguments)
    $baseArguments = @($entryArguments | Where-Object { $_ -notlike '--selftest-case=*' })
    $originalCaseFilter = ''
    foreach ($argument in $entryArguments) {
        if ($argument -like '--selftest-case=*') {
            $originalCaseFilter = $argument.Substring('--selftest-case='.Length)
            break
        }
    }

    $failedCaseNames = @($FailureNames | Where-Object {
            -not [string]::IsNullOrWhiteSpace($_) -and $_ -notin @('selftest_result_coverage', 'setup')
        } | Sort-Object -Unique)

    $plans = @()
    foreach ($caseName in $failedCaseNames) {
        $plans += [pscustomobject]@{
            name = $caseName
            mode = 'failed-case'
            arguments = @($baseArguments) + @("--selftest-case=$caseName")
            shuffle_seed = ''
        }
    }

    if (@($failedCaseNames).Count -eq 0) {
        return $plans
    }

    $shuffleBaseArguments = @($baseArguments | Where-Object {
            $_ -notlike '--selftest-shuffle=*' -and $_ -notlike '--selftest-repeat=*'
        })
    if (-not [string]::IsNullOrWhiteSpace($originalCaseFilter)) {
        $shuffleBaseArguments += "--selftest-case=$originalCaseFilter"
    }

    foreach ($seed in @($ShuffleTriageSeeds)) {
        if ([string]::IsNullOrWhiteSpace($seed)) {
            continue
        }

        $plans += [pscustomobject]@{
            name = [string](Get-RSObjectValue -Object $Entry -Names @('Name', 'name') -DefaultValue '')
            mode = 'shuffle-triage'
            arguments = @($shuffleBaseArguments) + @("--selftest-shuffle=$seed")
            shuffle_seed = $seed
        }
    }

    return $plans
}

Export-ModuleMember -Function @(
    'Get-RSSelfTestArguments',
    'Get-RSSelfTestListArguments',
    'New-RSPesterInvokeParameters',
    'Get-RSBuildScriptArguments',
    'Get-RSBuildEnvironmentOverrides',
    'Test-RSSelfTestResultCoverage',
    'Get-RSObjectValue',
    'Test-RSTestResultPassed',
    'Add-RSTestResultClassification',
    'Get-RSSelfTestClassificationRetryPlan'
)
