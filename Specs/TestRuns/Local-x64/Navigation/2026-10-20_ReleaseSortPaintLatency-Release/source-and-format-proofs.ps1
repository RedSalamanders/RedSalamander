$ErrorActionPreference='Stop'
$root=$PSScriptRoot
$repo='Z:/src/RedSalamander'
$qualified='Z:/src/RedSalamander-worktrees/debug-sort-row-moves-validation'
$oldRoot='Z:/RedSalamander.Perf/evidence/natural-sort-row-reorder-2026-10-20'
$old=Get-Content (Join-Path $oldRoot 'ordering-preservation-source-checks.json') -Raw|ConvertFrom-Json
$checks=@()
foreach($relative in @('Common/NaturalSort.h','RedSalamander/FolderViewSortPolicy.h')){
 $expected=@($old.protected_inputs|Where-Object path -eq $relative)
 $hash=(Get-FileHash -LiteralPath (Join-Path $repo $relative)).Hash
 if($expected.Count-ne1-or$expected[0].sha256-ne$hash-or(Get-FileHash -LiteralPath (Join-Path $qualified $relative)).Hash-ne$hash){throw 'Windows comparison/threshold input differs'}
 $checks += [ordered]@{path=$relative;sha256=$hash;unchanged=$true}
}
$initial=Get-Content (Join-Path $root 'initial-baseline-identities.json') -Raw|ConvertFrom-Json
$expected=@($initial.sources|Where-Object path -eq 'RedSalamander/FolderView.Enumeration.cpp')
$hash=(Get-FileHash -LiteralPath (Join-Path $repo 'RedSalamander/FolderView.Enumeration.cpp')).Hash
if($expected.Count-ne1-or$hash-ne$expected[0].sha256){throw 'Row ordering implementation changed'}
$checks += [ordered]@{path='RedSalamander/FolderView.Enumeration.cpp';sha256=$hash;unchanged=$true}
$renderPath='RedSalamander/FolderView.Rendering.cpp'
$before=(& git -C $repo show "d7dd952d629985b48d25b23b56d1726fbb2a1ec5:$renderPath")-join"`n"
$after=[IO.File]::ReadAllText((Join-Path $repo $renderPath)).Replace("`r`n","`n")
$pattern='#ifdef ENABLE_TESTS\n        if \(hrPresent == S_OK\)\n        \{\n            \+\+_debugSuccessfulPresentCount;\n        \}\n#endif\n'
if([regex]::Matches($after,$pattern).Count-ne2){throw 'Expected two test-only presentation counters'}
$stripped=[regex]::Replace($after,$pattern,'')
if(![StringComparer]::Ordinal.Equals($before.TrimEnd("`n"),$stripped.TrimEnd("`n"))){throw 'Production rendering policy changed beyond test-only observation'}
$before=(&git -C $repo show 'd7dd952d629985b48d25b23b56d1726fbb2a1ec5:RedSalamander/FolderView.cpp')-join"`n"
$after=[IO.File]::ReadAllText((Join-Path $repo 'RedSalamander/FolderView.cpp')).Replace("`r`n","`n")
$start='void FolderView::SetSort(';$end='void FolderView::SetShowHiddenFiles('
$a=$before.Substring($before.IndexOf($start),$before.IndexOf($end)-$before.IndexOf($start))
$b=$after.Substring($after.IndexOf($start),$after.IndexOf($end)-$after.IndexOf($start))
if(![StringComparer]::Ordinal.Equals($a,$b)){throw 'Product SetSort changed'}
$files=@('RedSalamander/FolderView.cpp','RedSalamander/FolderView.h','RedSalamander/FolderView.Enumeration.cpp','RedSalamander/FolderView.Rendering.cpp','RedSalamander/SelfTest/Commands/Commands.Tests.ViewCommands.cpp')
$source=@(foreach($relative in $files){
 $hash=(Get-FileHash -LiteralPath (Join-Path $repo $relative)).Hash
 if((Get-FileHash -LiteralPath (Join-Path $qualified $relative)).Hash-ne$hash){throw 'Main and qualified C++ differ'}
 [ordered]@{path=$relative;sha256=$hash;matches_qualified_checkout=$true}
})
$vs=& "${env:ProgramFiles(x86)}\Microsoft Visual Studio\Installer\vswhere.exe" -latest -prerelease -products '*' -property installationPath
$formatter=Join-Path $vs 'VC/Tools/Llvm/x64/bin/clang-format.exe'
$format=@(foreach($relative in $files){
 $ranges=@(&git -C $repo diff --unified=0 -- $relative|ForEach-Object {if($_-match'^@@ -\d+(?:,\d+)? \+(\d+)(?:,(\d+))? @@'){$start=[int]$Matches[1];$count=if($Matches[2]){[int]$Matches[2]}else{1};if($count){'-lines={0}:{1}'-f$start,($start+$count-1)}}})
 [xml]$xml=(& $formatter -output-replacements-xml @ranges (Join-Path $repo $relative))-join"`n"
 if($LASTEXITCODE-or@($xml.replacements.replacement|Where-Object {$null-ne$_}).Count){throw 'Changed-range formatting differs'}
 [ordered]@{path=$relative;replacements=0;ranges=$ranges}
})
&git -C $repo diff --check
if($LASTEXITCODE){throw 'Whitespace check failed'}
[ordered]@{verified_utc=[DateTime]::UtcNow.ToString('o');status='passed';unchanged_ordering_inputs=$checks;production_SetSort_unchanged_with_lf_normalization=$true;production_rendering_unchanged_after_removing_only_two_test_counter_blocks=$true;cpp_source=$source;formatter=$formatter;format=$format;scope='Production comparator, reorder implementation, SetSort and rendering policy remain unchanged. Test-only observation and benchmark protocol change. Fixture and request sequence are independently bound in ordering-and-fixture-checks.json.'}|ConvertTo-Json -Depth 9|Set-Content (Join-Path $root 'source-and-format-proofs.json')
[pscustomobject]@{status='passed';matching_cpp_files=$source.Count;format_checks=$format.Count}
