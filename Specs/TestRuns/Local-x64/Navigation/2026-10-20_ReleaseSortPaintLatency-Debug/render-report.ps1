$ErrorActionPreference='Stop'
$root=$PSScriptRoot
$base='Z:/src/RedSalamander/Specs/TestRuns/Local-x64/Navigation'
$culture=[Globalization.CultureInfo]::InvariantCulture
function Ms($v){([decimal]$v/1000).ToString('0.000',$culture)}
$summaries=@{}
foreach($configuration in @('Debug','Release')){$summaries[$configuration]=Get-Content (Join-Path $base "2026-10-20_ReleaseSortPaintLatency-$configuration/performance-comparison.json") -Raw|ConvertFrom-Json}
foreach($configuration in @('Debug','Release')){
 $s=$summaries[$configuration]
 $packet=Join-Path $base "2026-10-20_ReleaseSortPaintLatency-$configuration"
 $text=[Collections.Generic.List[string]]::new()
 foreach($line in @("# I44 Release sort paint latency - $configuration",'',
 'The apparent Extension Ascending regression was dominated by presentation waits in a synthetic warm-render sequence. The old recorder timed the product paint plus an injected invalidation, swap-chain resize/deferred initialization, icon work, message draining and extra renders. Its four-frame requests were not the same quantity as one product sort paint.',
 '',
 'The corrected `folderView_perf_sort_toggle_stress` prepares and settles outside timing, then observes the product paint through `Present == S_OK`, without synthetic warming or invalidation inside the measured request. Failed EndDraw/Present and occluded frames cannot satisfy it. A warm no-op does not manufacture a frame. The legacy control retains its original sequence and metric name.',
 '',
 '**This is a benchmark protocol correction, not a production speedup.** Windows comparison, the full row-ordering implementation, SetSort, rendering policy, VSync, the 5,000-file corpus and ten requests are unchanged. See the source/fixture checks. The test-only counters add observation, not production behavior.',
 '',
 '## Counterbalanced same-executable results',
 '',
 'Milliseconds, median [minimum - maximum]. Each protocol has two five-repeat blocks: ten observations per request. Debug runs legacy/product/product/legacy; Release runs product/legacy/legacy/product. The two protocols use one exact candidate executable receipt per profile. Different metric definitions are not pooled or converted into a product speedup percentage.',
 '',
 '| Request | Product paint v2 ms | Legacy warm v1 ms |',
 '| --- | --- | --- |')){$text.Add($line)}
 for($i=0;$i-lt10;$i++){
  $p=$s.role_statistics.'candidate product paint'[$i];$l=$s.role_statistics.'candidate legacy control'[$i]
  $label="$($i+1). $($p.key) $($p.direction)"
  $text.Add("| $label | $(Ms $p.stats.median_us) [$(Ms $p.stats.min_us) - $(Ms $p.stats.max_us)] | $(Ms $l.stats.median_us) [$(Ms $l.stats.min_us) - $(Ms $l.stats.max_us)] |")
 }
 $productPhases=@($s.phases.PSObject.Properties.Value|Where-Object {$_.selected-and$_.role-eq'candidate product paint'})
 $productSamples=@($productPhases|ForEach-Object samples)
 $legacyPhases=@($s.phases.PSObject.Properties.Value|Where-Object {$_.selected-and$_.role-eq'candidate legacy control'})
 $legacySamples=@($legacyPhases|ForEach-Object samples)
 $text.Add('');$text.Add("Across $($productSamples.Count) product requests, synthetic warm calls total $(($productSamples|Measure-Object -Property warm_calls -Sum).Sum); renders total $(($productSamples|Measure-Object -Property render_calls -Sum).Sum); successful presents total $(($productSamples|Measure-Object -Property successful_presents -Sum).Sum). Every changed sort has a successful paint; all ten no-op requests have zero renders. The $($legacySamples.Count) legacy-control requests retain one warm call each. All per-request counters remain in performance-comparison.json and original metric rows.")
 foreach($line in @('',
 ('The initial unchanged-code baseline used the earlier build-only observer. It remains retained as an informational old-protocol reference; it does not establish absence of concurrent native test runtimes. Its Extension Ascending median/range is '+(Ms $s.role_statistics.'fresh legacy baseline'[1].stats.median_us)+' ms ['+(Ms $s.role_statistics.'fresh legacy baseline'[1].stats.min_us)+' - '+(Ms $s.role_statistics.'fresh legacy baseline'[1].stats.max_us)+' ms]. It is not substituted for the stronger same-executable counterbalanced controls.'),
 '',
 '## Runs, receipts and observation',
 '',
 '| Phase | Selected | Run ID | Receipt | Overlap polls | Max gap ms | Observer CPU ms |',
 '| --- | --- | --- | --- | --- | --- | --- |')){$text.Add($line)}
 foreach($entry in $s.phases.PSObject.Properties){$p=$entry.Value;$text.Add("| $($entry.Name) | $($p.selected) | $($p.run_id) | $($p.receipt_id) | $($p.build_overlap_polls) | $($p.maximum_poll_gap_ms) | $($p.observer_cpu_ms) |")}
 foreach($line in @('',
 'The final v2 observer requires a 30-second quiet gate and monitors the whole canonical runner window, nominally every 200 ms. It detects compilers/build launchers, known DxUi functional/performance executables and test-profile drivers. Exact old idle MSBuild identities require dead-parent and unchanged-CPU verification; resumed CPU is a blocker. The unrelated verified GIGABYTE GCC process has a documented exact identity exclusion.',
 '',
 'This is snapshot-based absence of detected build/known-test overlap, not a lossless process-start audit, proof of an idle host, or physical scanout measurement. Present success means the API returned S_OK. Unknown shells, actual gaps and observer CPU are retained. No independent application, build or test was interrupted.',
 '',
 'A gate-only v1 controller was retired before any native runner launched so the known concurrent DxUi test profile could become a blocker. A subsequent block passed natively but was conservatively rejected when a metadata-preflight shell mentioning test.ps1 was observed. The metadata run itself was deferred. Every rejected measurement and the retirement remain separate; acceptance never selects by latency.',
 '',
 '## Correctness and scope',
 '',
 'Both isolated test-enabled x64 Debug/Release solution builds finish with zero warnings/errors. Nine native cases pass per profile: Windows natural-order/reorder contracts, five current/selection regressions, successful product-paint contracts, rendering-error persistence and device-loss recovery. The new paint contract covers changed/no-op sorts, invalid targets, failed EndDraw/Present, occlusion and successful later paint without synthetic warming in observation. All 328 focused tooling/source-contract tests pass on the qualified source; final main metadata validation is recorded separately.',
 '',
 'All original ordering/cache/numeric-metadata rules remain in force. No production render scheduling, VSync or comparator optimization is claimed. This does not establish Full, ARM64, ASan or completely idle-host qualification. The shared sandbox audit retains 41 existing unexpected run-directory findings; native case passes do not make that audit clean.',
 '',
 '## Evidence retention and reproduction',
 '',
 'Every request sample and relevant sorting/layout/presentation metric row is retained in original order. Complete measurement snapshots and the final gate interval, original verdicts/traces/receipts and exact extraction rules are bound in artifact-bindings.json. Larger original streams and full build logs remain under `Z:/RedSalamander.Perf/evidence/natural-sort-release-extension-2026-10-20`; original/extract digests and byte/row counts are retained. Receipt JSON is deduplicated by immutable receipt ID, with an explicit reference per phase.',
 '',
 'The observer detects the live concurrent DxUi Accessibility process in all five positive-control observations. The generic probe label says synthetic; probe-context.json records that the already-running real test was observed and no test was launched. The two final observer scripts differ only in case/artifact selectors. Initial v1 and final v2 source/scanner identities stay distinct.',
 '',
 'The original I43 packet remains immutable. A separately retained Debug smoke pass overlaps the Release build and is functional evidence only. Core-ordering records include setup/restoration; separate nested stage medians are not additive. Full raw native metric streams, including enumeration/preparation records, remain digest-bound.',
 '',
 'Canonical commands for each profile:',
 '',
 '```powershell',
 "Tools/Run-AllTests.ps1 -Suite Commands -Configuration $configuration -SkipBuild -CaseFilter folderView_perf_sort_toggle_stress -SelfTestRepeat 5",
 "Tools/Run-AllTests.ps1 -Suite Commands -Configuration $configuration -SkipBuild -CaseFilter folderView_perf_sort_toggle_legacy_warm_control -SelfTestRepeat 5",
 '```',
 '',
 'The archived orchestration scripts are evidence snapshots, not new supported repository commands. The qualified checkout is `Z:/src/RedSalamander-worktrees/debug-sort-row-moves-validation`; main contains the same five C++ inputs, while the independently running main Release instance remains undisturbed.',
 '',
 "[Other profile](../2026-10-20_ReleaseSortPaintLatency-$(if($configuration-eq'Debug'){'Release'}else{'Debug'})/README.md) · [Independent verification](independent-verification.json)")){$text.Add($line)}
 [IO.File]::WriteAllLines((Join-Path $packet 'README.md'),$text.ToArray(),[Text.UTF8Encoding]::new($false))
}
