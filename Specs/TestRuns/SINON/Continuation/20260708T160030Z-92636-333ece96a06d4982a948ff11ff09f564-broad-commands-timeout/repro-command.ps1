$env:REDSALAMANDER_TEST_ROOT='C:\RSPerf'
Remove-Item Env:\REDSALAMANDER_TEST_RUN_ID -ErrorAction SilentlyContinue
.\Tools\Run-AllTests.ps1 -Suite Commands -SkipBuild -SkipLegacySandboxCleanup -TimeoutMultiplier 2
