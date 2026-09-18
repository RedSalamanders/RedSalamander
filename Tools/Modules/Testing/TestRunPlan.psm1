Set-StrictMode -Version Latest

# Compatibility facade for the established TestRunPlan module surface.
$domainModules = @(
    'TestRunPlan.Common.psm1',
    'TestSandbox.psm1',
    'TestMachineResources.psm1',
    'TestInvocation.psm1',
    'TestQuarantine.psm1',
    'TestRunSummary.psm1',
    'TestSuitePlan.psm1'
)
foreach ($domainModule in $domainModules) {
    Import-Module (Join-Path $PSScriptRoot $domainModule) -Force -Scope Local
}

Export-ModuleMember -Function @(
    'New-RSTestRunPlanEntry',
    'Get-RSPesterEnvironmentBoundSkipCases',
    'Assert-RSTestRunPlanEntries',
    'ConvertTo-RSValidationPlanEntry',
    'New-RSValidationPlan',
    'Get-RSBuildOutputPath',
    'ConvertTo-RSFullPath',
    'New-RSTestRunId',
    'ConvertFrom-RSTestRunId',
    'Get-RSLiveProcessSnapshots',
    'Initialize-RSTestSandboxRoot',
    'Assert-RSTestSandboxContainedPath',
    'Get-RSTestSandboxRoot',
    'Get-RSTestMachineResourceManifestPath',
    'Read-RSTestMachineResources',
    'Set-RSTestMachineResourceEnvironment',
    'Restore-RSTestMachineResourceEnvironment',
    'New-RSTestRunContext',
    'New-RSTestSandboxScratchDirectory',
    'Get-RSToolingPesterSandboxRoot',
    'Remove-RSTestSandboxExactRunDirectory',
    'Resolve-RSTestSandboxStaleRunTargets',
    'Remove-RSTestSandboxStaleRunDirectories',
    'Get-RSTestSandboxDiskAudit',
    'Get-RSSelfTestArtifactRoot',
    'Get-RSSelfTestArguments',
    'Get-RSSelfTestListArguments',
    'New-RSPesterInvokeParameters',
    'Get-RSBuildScriptArguments',
    'Get-RSBuildEnvironmentOverrides',
    'Test-RSSelfTestResultCoverage',
    'Get-RSObjectValue',
    'Test-RSTestResultPassed',
    'Add-RSTestResultClassification',
    'Get-RSSelfTestClassificationRetryPlan',
    'Get-RSTestQuarantineDate',
    'Get-RSTestQuarantineStatus',
    'Read-RSTestQuarantineFile',
    'Get-RSTestQuarantineRepairPlan',
    'Convert-RSTestCaseForRunSummary',
    'Convert-RSTestRetryAttemptForRunSummary',
    'Convert-RSTestQuarantineRepairAttemptForRunSummary',
    'Get-RSValidationExitCode',
    'New-RSTestRunSummary',
    'New-RSValidationRunSummaryV2',
    'Convert-RSTestRunSummaryToCaseHistoryRows',
    'Convert-RSTestCaseHistoryRowsToJsonl',
    'Convert-RSTestCaseHistoryRowsToDashboardMarkdown',
    'Convert-RSTestRunSummaryToGitHubStepSummary',
    'Get-RSCommandsSelfTestFamilies',
    'Get-RSTestRunPlan'
)
