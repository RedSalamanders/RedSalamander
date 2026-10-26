Set-Location 'Z:\src\RedSalamander'
.\Tools\Run-AllTests.ps1 -Suite Full -TimeoutMultiplier 2 *> 'Z:\src\RedSalamander\.build\codex-runs\full_after_dxui_viewerpe_green_20260704.out.log'
$LASTEXITCODE | Set-Content -Path 'Z:\src\RedSalamander\.build\codex-runs\full_after_dxui_viewerpe_green_20260704.exit.txt'
