Describe 'Native build console and short dependency paths' {
    It 'refuses warnings and collisions, emits safe links and never auto-runs output' {
        & (Join-Path $PSScriptRoot 'BuildConsoleCheck.ps1')
    }
}
Describe 'Build console fixture run lifetime' {
    BeforeAll {
        $script:buildConsoleRepo=[IO.Path]::GetFullPath((Join-Path $PSScriptRoot '../..'))
        $script:buildConsoleHelper=Join-Path $PSScriptRoot 'BuildConsoleCheck.ps1'
        Import-Module (Join-Path $script:buildConsoleRepo 'Tools/Modules/Testing/TestSandbox.psm1')
        Import-Module (Join-Path $script:buildConsoleRepo 'Tools/Modules/Testing/TestRunPlan.Common.psm1')
        # Exercise the real lifecycle around a sentinel/failing body; the original
        # end-to-end test above still executes every console/path assertion.
        $text=[IO.File]::ReadAllText($script:buildConsoleHelper)
        $bodyStart=$text.IndexOf('Import-Module (Join-Path $Root ''Build/BuildConsole.psm1'')',[StringComparison]::Ordinal)
        $bodyEnd=$text.IndexOf('} catch {$failure=$_} finally {',[StringComparison]::Ordinal)
        if($bodyStart -lt 0 -or $bodyEnd -le $bodyStart){throw 'BuildConsole lifecycle boundaries changed; review this test.'}
        $script:buildConsoleBoundary=[scriptblock]::Create($text.Substring(0,$bodyStart)+'& $script:buildConsoleLifecycleBody'+[Environment]::NewLine+$text.Substring($bodyEnd))
    }

    It 'uses the current canonical module closure for a representative bp28 shared fixture' {
        # This is a current representative copy, not evidence about an unknown
        # historical shared fixture. Common is already a TestSandbox dependency.
        $fixture=Join-Path $TestDrive 'canonical-closure/bp28'
        [void][IO.Directory]::CreateDirectory((Join-Path $fixture 'Tools'))
        [void][IO.Directory]::CreateDirectory((Join-Path $fixture 'Build'))
        Copy-Item -LiteralPath (Join-Path $script:buildConsoleRepo 'Tools/Modules') -Destination (Join-Path $fixture 'Tools/Modules') -Recurse
        Copy-Item -LiteralPath (Join-Path $script:buildConsoleRepo 'Build/BuildConsole.psm1') -Destination (Join-Path $fixture 'Build/BuildConsole.psm1')
        Copy-Item -LiteralPath (Join-Path $script:buildConsoleRepo 'Tools/Modules/Build/DxUiPaths.psm1') -Destination (Join-Path $fixture 'Build/DxUiPaths.psm1')
        (Test-Path -LiteralPath (Join-Path $fixture 'Tools/Modules/Testing/TestRunPlan.Common.psm1')) | Should Be $true
        & $script:buildConsoleHelper -Root $fixture
    }

    It 'preserves explicit scratch compatibility without borrowing an invalid run id' {
        $fixture=Join-Path $TestDrive 'explicit-fixture'
        [void][IO.Directory]::CreateDirectory((Join-Path $fixture 'Build'))
        Copy-Item -LiteralPath (Join-Path $script:buildConsoleRepo 'Build/BuildConsole.psm1') -Destination (Join-Path $fixture 'Build/BuildConsole.psm1')
        Copy-Item -LiteralPath (Join-Path $script:buildConsoleRepo 'Tools/Modules/Build/DxUiPaths.psm1') -Destination (Join-Path $fixture 'Build/DxUiPaths.psm1')
        $previous=$env:REDSALAMANDER_TEST_RUN_ID
        try {
            $env:REDSALAMANDER_TEST_RUN_ID='console-invalid-explicit-fixture'
            & $script:buildConsoleHelper -Root $fixture -ScratchRoot (Join-Path $fixture 'scratch')
        } finally {$env:REDSALAMANDER_TEST_RUN_ID=$previous}
    }

    Context 'borrowed and standalone execution boundaries' {
        BeforeEach {
            $script:buildConsolePreviousRunId=$env:REDSALAMANDER_TEST_RUN_ID
            $script:buildConsoleTestRoot=Join-Path $TestDrive ([guid]::NewGuid().ToString('N'))
            $runId=New-RSTestRunId
            $script:buildConsoleContext=[pscustomobject]@{TestRoot=$script:buildConsoleTestRoot;RunId=$runId;RunRoot=(Join-Path $script:buildConsoleTestRoot ('runs/'+$runId))}
            [void][IO.Directory]::CreateDirectory($script:buildConsoleContext.RunRoot)
            $script:buildConsoleLifecycleBody={$script:buildConsoleBodyReached=$true}
            $script:buildConsoleBodyReached=$false
            # Actual canonical modules were imported above. Block reimport only
            # inside this recording boundary so mocks cannot be overwritten.
            Mock Import-Module {}
            Mock Get-RSTestSandboxRoot {return $script:buildConsoleTestRoot}
            Mock Assert-RSTestSandboxContainedPath {
                param($TestRoot,$Path)
                if(-not $Path.StartsWith($TestRoot+'\',[StringComparison]::OrdinalIgnoreCase)){throw 'Fixture containment refusal.'}
                return $Path
            }
            Mock New-RSTestRunContext {return $script:buildConsoleContext}
            Mock New-RSTestSandboxScratchDirectory {return (Join-Path $script:buildConsoleContext.RunRoot 'scratch/BuildConsole/Policy')}
            Mock Save-RSTestSandboxRunEvidence {return [pscustomobject]@{Status='Archived';RunRoot=(Join-Path $script:buildConsoleTestRoot 'archive')}}
            Mock Write-Host {}
        }
        AfterEach {$env:REDSALAMANDER_TEST_RUN_ID=$script:buildConsolePreviousRunId}

        It 'borrows an existing canonical owner without archiving the caller context' {
            $env:REDSALAMANDER_TEST_RUN_ID=$script:buildConsoleContext.RunId
            & $script:buildConsoleBoundary -Root $script:buildConsoleRepo
            $script:buildConsoleBodyReached | Should Be $true
            Assert-MockCalled -Scope It New-RSTestRunContext -Times 1 -Exactly
            Assert-MockCalled -Scope It Save-RSTestSandboxRunEvidence -Times 0 -Exactly
            [IO.Directory]::Exists($script:buildConsoleContext.RunRoot) | Should Be $true
        }

        It 'refuses a missing borrowed run before a context factory that would create it' {
            $runId=New-RSTestRunId
            $script:buildConsoleContext.RunId=$runId
            $script:buildConsoleContext.RunRoot=Join-Path $script:buildConsoleTestRoot ('runs/'+$runId)
            $env:REDSALAMANDER_TEST_RUN_ID=$runId
            Mock New-RSTestRunContext {
                [void][IO.Directory]::CreateDirectory($script:buildConsoleContext.RunRoot)
                return $script:buildConsoleContext
            }
            {& $script:buildConsoleBoundary -Root $script:buildConsoleRepo} | Should Throw 'borrowed run root does not already exist'
            Assert-MockCalled -Scope It New-RSTestRunContext -Times 0 -Exactly
            Assert-MockCalled -Scope It New-RSTestSandboxScratchDirectory -Times 0 -Exactly
            [IO.Directory]::Exists($script:buildConsoleContext.RunRoot) | Should Be $false
        }

        It 'rejects <Kind> ids before authorizing or creating a borrowed context' -TestCases @(
            @{Kind='malformed'}
            @{Kind='native fallback'}
            @{Kind='harness fallback'}
        ) {
            param($Kind)
            $env:REDSALAMANDER_TEST_RUN_ID=switch($Kind){
                'malformed' {'console-'+[guid]::NewGuid().ToString('N')}
                'native fallback' {'adhoc-'+$PID}
                'harness fallback' {'harness-'+$PID+'-1234'}
            }
            {& $script:buildConsoleBoundary -Root $script:buildConsoleRepo} | Should Throw 'canonical active runner id'
            Assert-MockCalled -Scope It Get-RSTestSandboxRoot -Times 0 -Exactly
            Assert-MockCalled -Scope It New-RSTestRunContext -Times 0 -Exactly
            Assert-MockCalled -Scope It New-RSTestSandboxScratchDirectory -Times 0 -Exactly
        }

        It 'refuses an existing run whose owner is no longer live' {
            $runId=[DateTime]::UtcNow.ToString('yyyyMMddTHHmmssZ')+'-2147483647-'+[guid]::NewGuid().ToString('N')
            $script:buildConsoleContext.RunId=$runId
            $script:buildConsoleContext.RunRoot=Join-Path $script:buildConsoleTestRoot ('runs/'+$runId)
            [void][IO.Directory]::CreateDirectory($script:buildConsoleContext.RunRoot)
            $env:REDSALAMANDER_TEST_RUN_ID=$runId
            # Pester 3.4's bare Should Throw reports false for an empty expected
            # message. Inspect the real GetProcessById refusal and its identity.
            $deadOwnerError=$null
            try {& $script:buildConsoleBoundary -Root $script:buildConsoleRepo} catch {$deadOwnerError=$_}
            ($null -ne $deadOwnerError) | Should Be $true
            $deadOwnerCause=$deadOwnerError.Exception
            while($null -ne $deadOwnerCause.InnerException){$deadOwnerCause=$deadOwnerCause.InnerException}
            ($deadOwnerCause -is [ArgumentException]) | Should Be $true
            $deadOwnerCause.Message | Should Match '2147483647'
            $script:buildConsoleBodyReached | Should Be $false
            Assert-MockCalled -Scope It Get-RSTestSandboxRoot -Times 1 -Exactly
            Assert-MockCalled -Scope It Assert-RSTestSandboxContainedPath -Times 1 -Exactly
            Assert-MockCalled -Scope It New-RSTestRunContext -Times 0 -Exactly
            Assert-MockCalled -Scope It New-RSTestSandboxScratchDirectory -Times 0 -Exactly
        }

        It 'refuses an existing run whose owner PID has been reused' {
            [IO.Directory]::SetCreationTimeUtc($script:buildConsoleContext.RunRoot,[DateTime]::new(2000,1,1))
            $env:REDSALAMANDER_TEST_RUN_ID=$script:buildConsoleContext.RunId
            {& $script:buildConsoleBoundary -Root $script:buildConsoleRepo} | Should Throw 'inactive or its PID was reused'
            Assert-MockCalled -Scope It New-RSTestRunContext -Times 0 -Exactly
            Assert-MockCalled -Scope It New-RSTestSandboxScratchDirectory -Times 0 -Exactly
        }

        It 'fails <Boundary> authorization before the context factory' -TestCases @(
            @{Boundary='marker'}
            @{Boundary='containment'}
        ) {
            param($Boundary)
            $env:REDSALAMANDER_TEST_RUN_ID=$script:buildConsoleContext.RunId
            if($Boundary -ceq 'marker'){Mock Get-RSTestSandboxRoot {throw 'Fixture authorization refusal.'}}
            else{Mock Assert-RSTestSandboxContainedPath {throw 'Fixture authorization refusal.'}}
            {& $script:buildConsoleBoundary -Root $script:buildConsoleRepo} | Should Throw 'Fixture authorization refusal'
            Assert-MockCalled -Scope It New-RSTestRunContext -Times 0 -Exactly
            Assert-MockCalled -Scope It New-RSTestSandboxScratchDirectory -Times 0 -Exactly
        }

        It 'archives only its standalone canonical context after the body succeeds' {
            $env:REDSALAMANDER_TEST_RUN_ID=$null
            & $script:buildConsoleBoundary -Root $script:buildConsoleRepo
            Assert-MockCalled -Scope It Save-RSTestSandboxRunEvidence -Times 1 -Exactly -ParameterFilter {$RunId -ceq $script:buildConsoleContext.RunId -and $Disposition -ceq 'completed'}
            Assert-MockCalled -Scope It Write-Host -Times 1 -Exactly -ParameterFilter {$Object -like 'PASS zero-warning*'}
        }

        It 'retains primary and archive errors including a fatal descendant for <Name>' -TestCases @(
            @{Name='ordinary body and archive';PrimaryFatal=$false;ArchiveFatal=$false}
            @{Name='ordinary body and fatal archive';PrimaryFatal=$false;ArchiveFatal=$true}
            @{Name='fatal body and ordinary archive';PrimaryFatal=$true;ArchiveFatal=$false}
        ) {
            param($Name,$PrimaryFatal,$ArchiveFatal)
            $env:REDSALAMANDER_TEST_RUN_ID=$null
            $script:buildConsolePrimary=if($PrimaryFatal){[OutOfMemoryException]::new('Body fatal')}else{[InvalidOperationException]::new('Body failure')}
            $script:buildConsoleArchiveError=if($ArchiveFatal){[OutOfMemoryException]::new('Archive fatal')}else{[IO.IOException]::new('Archive failure')}
            $script:buildConsoleLifecycleBody={throw $script:buildConsolePrimary}
            Mock Save-RSTestSandboxRunEvidence {throw $script:buildConsoleArchiveError}
            $captured=$null
            try{& $script:buildConsoleBoundary -Root $script:buildConsoleRepo}catch{$captured=$_}
            ($null -ne $captured) | Should Be $true
            $cause=$captured.Exception
            while($cause -isnot [AggregateException] -and $null -ne $cause.InnerException){$cause=$cause.InnerException}
            ($cause -is [AggregateException]) | Should Be $true
            $cause.InnerExceptions.Count | Should Be 2
            @($cause.InnerExceptions|Where-Object {[object]::ReferenceEquals($_,$script:buildConsolePrimary)}).Count | Should Be 1
            @($cause.InnerExceptions|Where-Object {[object]::ReferenceEquals($_,$script:buildConsoleArchiveError)}).Count | Should Be 1
            Assert-MockCalled -Scope It Save-RSTestSandboxRunEvidence -Times 1 -Exactly -ParameterFilter {$Disposition -ceq 'interrupted'}
            Assert-MockCalled -Scope It Write-Host -Times 0 -Exactly -ParameterFilter {$Object -like 'PASS zero-warning*'}
        }

        It 'does not print success if the standalone archive is retained rather than archived' {
            $env:REDSALAMANDER_TEST_RUN_ID=$null
            Mock Save-RSTestSandboxRunEvidence {return [pscustomobject]@{Status='Retained';RunRoot=$script:buildConsoleContext.RunRoot}}
            {& $script:buildConsoleBoundary -Root $script:buildConsoleRepo} | Should Throw 'standalone evidence was not archived'
            Assert-MockCalled -Scope It Write-Host -Times 0 -Exactly -ParameterFilter {$Object -like 'PASS zero-warning*'}
        }
    }
}
