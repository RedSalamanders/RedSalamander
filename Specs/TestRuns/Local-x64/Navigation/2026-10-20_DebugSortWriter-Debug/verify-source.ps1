$ErrorActionPreference='Stop'
$root=$PSScriptRoot
$main='Z:/src/RedSalamander'
$baseline='E:/RedSalamander.SortValidation-20261020/debug-sort-paint-bottleneck'
$candidate='S:/RedSalamander.SortValidation-20261020/debug-sort-writer-validation'
$protected=@('Common/NaturalSort.h','Common/HandleIo.h','RedSalamander/FolderViewSortPolicy.h','RedSalamander/FolderView.Enumeration.cpp','RedSalamander/FolderView.Layout.cpp','RedSalamander/FolderView.Rendering.cpp','RedSalamander/FolderView.cpp','RedSalamander/FolderView.h','Directory.Build.props','Directory.Build.targets','Dependencies/DxUi.lock.json')
$proof=@(foreach($relative in $protected){
 &git -C $main diff --quiet a6ced5f41 -- $relative
 if($LASTEXITCODE){throw "Protected product behavior changed: $relative"}
 $hash=(Get-FileHash -LiteralPath (Join-Path $main $relative)).Hash
 foreach($repo in @($baseline,$candidate)){if((Get-FileHash -LiteralPath (Join-Path $repo $relative)).Hash-ne$hash){throw "Qualified protected input differs: $relative"}}
 [ordered]@{path=$relative;sha256=$hash;unchanged=$true;all_three_checkouts_match=$true}
})
$compiled=@(foreach($relative in @('Common/Common/PerfJsonl.cpp','Common/Common.Tests.PerfJsonl.h','RedSalamander/SelfTest/Commands/Commands.Tests.Runner.cpp','RedSalamander/SelfTest/Commands/Commands.Tests.ViewCommands.cpp')){
 $hash=(Get-FileHash -LiteralPath (Join-Path $main $relative)).Hash
 if((Get-FileHash -LiteralPath (Join-Path $candidate $relative)).Hash-ne$hash){throw "Main candidate source mismatch: $relative"}
 if($relative-ne'Common/Common/PerfJsonl.cpp'-and(Get-FileHash -LiteralPath (Join-Path $baseline $relative)).Hash-ne$hash){throw 'Benchmark or diagnostic interface changed between baseline and candidate'}
 [ordered]@{path=$relative;sha256=$hash;main_matches_candidate=$true;baseline_matches=($relative-ne'Common/Common/PerfJsonl.cpp')}
})
$before=(&git -C $main show 'a6ced5f41:Common/Common/PerfJsonl.cpp')-join"`n"
$instrumented=[IO.File]::ReadAllText((Join-Path $baseline 'Common/Common/PerfJsonl.cpp')).Replace("`r`n","`n")
$stripped=$instrumented.Replace('#include "Common.Tests.PerfJsonl.h"'+"`n",'')
$pattern='(?ms)^#ifdef ENABLE_TESTS\n.*?^#endif\n?'
if([regex]::Matches($stripped,$pattern).Count-ne3){throw 'Expected exactly three baseline diagnostic blocks'}
$stripped=[regex]::Replace($stripped,$pattern,'')
$canonical={param($text)(($text-split"`n"|Where-Object {$_.Trim()})-join"`n").Trim()}
if((&$canonical $before)-cne(&$canonical $stripped)){throw 'Baseline production writer changed beyond test-only observation'}
$source=[IO.File]::ReadAllText((Join-Path $candidate 'Common/Common/PerfJsonl.cpp'))
$start=$source.IndexOf('COMMON_API JsonlDebugSnapshot GetJsonlDebugSnapshotForSelfTest()')
$end=$source.IndexOf('COMMON_API void ConfigureJsonlOutput(', $start)
$observer=$source.Substring($start,$end-$start)
if($observer-match'scoped_lock|unique_lock|lock_guard|WaitFor|Sleep\('){throw 'Candidate observer waits outside the measured span'}
[ordered]@{verified_utc=[DateTime]::UtcNow.ToString('o');status='passed';protected_inputs=$proof;compiled_inputs=$compiled;baseline_production_writer_unchanged_after_removing_exact_test_blocks=$true;benchmark_and_diagnostic_interface_identical=$true;candidate_observation='TLS plus always-lock-free atomic snapshots; no writer synchronization';scope='Only the Common JSONL writer is optimized. Windows ordering, row reordering, layout, drawing, presentation, compiler flags, pin, fixture, request sequence and timer boundaries are unchanged.'}|ConvertTo-Json -Depth 8|Set-Content (Join-Path $root 'source-preservation.json')
'Source preservation passed.'
