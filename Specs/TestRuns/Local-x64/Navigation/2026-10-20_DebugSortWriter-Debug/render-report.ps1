$ErrorActionPreference='Stop'
$root=$PSScriptRoot
$data=Get-Content (Join-Path $root 'performance-comparison.json') -Raw|ConvertFrom-Json
function Format-Ms($value){([double]$value/1000).ToString('F3',[Globalization.CultureInfo]::InvariantCulture)}
foreach($profile in @('Debug','Release')){
 $packet="Z:/src/RedSalamander/Specs/TestRuns/Local-x64/Navigation/2026-10-20_DebugSortWriter-$profile"
 New-Item -ItemType Directory -Path $packet -Force|Out-Null
 $text=[Collections.Generic.List[string]]::new()
 $text.Add("# I45 Debug sort writer optimization - $profile")
 $text.Add('')
 $text.Add('The unchanged 5,000-file product-paint scenario attributes this UI thread''s synchronous JSONL writer cost before and after the optimization. Debug enables this diagnostic path by default. The old writer copied and converted constant run metadata and formatted timestamp/numeric fields repeatedly, including one record for every drawn item.')
 $text.Add('')
 $text.Add('The candidate caches escaped metadata per configuration, proves short ASCII into bounded stack scratch, retains the strict Unicode fallback, appends unescaped runs in bulk, formats ASCII numbers/timestamps directly and reuses a row buffer with a 64-KiB retained ceiling. Configuration and append share one mutex. Appends remain synchronous and immediately readable; no metric is sampled, suppressed, deferred or removed. The shared partial-write helper preserves complete successful appends.')
 $text.Add('')
 $text.Add('Windows natural comparison, listing-local ranks, full row reordering, current/selection, Time/Size numeric keys, layout, drawing, presentation, VSync, checked Debug iterators and graphics validation are unchanged. This is a captured/default-Debug path optimization; it does not establish the same improvement with capture disabled. Both versions use identical test-only per-thread writer observations. Counter reads never acquire the writer lock or hide contention outside the timer. Writer time is inclusive elapsed time, including contention, not CPU time.')
 $text.Add('')
 $text.Add('The complete pairs localize the large Debug gain to recording: Extension Ascending paint median falls from 34.701 to 19.810 ms and counted writer spans from 17.312 to 2.479 ms. The per-request arithmetic remainder stays about 17 ms (17.467 to 17.205 ms). Debug ordering and drawing retain their ordinary unoptimized/checking costs: candidate OrderItems is 6.527 ms and begin-to-EndDraw 6.728 ms, versus Release 0.648/3.110 ms. These nested scopes include other diagnostic work and are not additive or pure CPU samples.')
 $text.Add('')
 $text.Add('Release Extension Ascending is a smaller median change, 8.984 to 8.591 ms, with overlapping ranges. Name Descending and the final None request have small higher candidate paint medians, also with overlapping ranges; no uniform Release speedup is claimed. The large Time Ascending change (24.106 to 7.945 ms) includes Present falling from 17.407 to 2.497 ms while writer spans fall only 0.959 to 0.568 ms. The data does not isolate why that presentation wait changed; it cannot establish a 16-ms CPU/comparison gain.')
 $text.Add('')
 $text.Add('## Same-protocol, counterbalanced results')
 $text.Add('')
 $text.Add('Milliseconds, median [minimum - maximum]. Each role has two five-repeat blocks and ten observations per request. Debug runs baseline/candidate/candidate/baseline; Release runs the reverse. Each role/profile has one exact executable receipt. The fixture, ten requests and product_paint_v2 timer boundaries are unchanged.')
 $text.Add('')
 $text.Add('| Request | Baseline paint ms | Candidate paint ms | Baseline writer ms | Candidate writer ms |')
 $text.Add('| --- | ---: | ---: | ---: | ---: |')
 foreach($index in 1..10){
  $b=@($data.statistics|Where-Object {$_.configuration-eq$profile-and$_.role-eq'baseline'-and$_.request_index-eq$index})[0]
  $c=@($data.statistics|Where-Object {$_.configuration-eq$profile-and$_.role-eq'candidate'-and$_.request_index-eq$index})[0]
  $text.Add("| $index. $($b.key) $($b.direction) | $(Format-Ms $b.paint.median_us) [$(Format-Ms $b.paint.min_us) - $(Format-Ms $b.paint.max_us)] | $(Format-Ms $c.paint.median_us) [$(Format-Ms $c.paint.min_us) - $(Format-Ms $c.paint.max_us)] | $(Format-Ms $b.writer.median_us) | $(Format-Ms $c.writer.median_us) |")
 }
 $text.Add('')
 $text.Add('Every request and its writer-call, render, successful-Present and synthetic-warm counters remain in performance-comparison.json and selected original JSONL rows. A no-op cannot manufacture a frame. Changed sorts require a successful S_OK presentation; failed/occluded attempts do not qualify. Separate nested stage distributions are not additive, and recording cost inside drawing/layout scopes is not pure graphics CPU time. Full originals retain all per-item draw and frame families.')
 $text.Add('')
 $text.Add('### Extension Ascending attribution')
 $text.Add('')
 $b=@($data.statistics|Where-Object {$_.configuration-eq$profile-and$_.role-eq'baseline'-and$_.request_index-eq2})[0]
 $c=@($data.statistics|Where-Object {$_.configuration-eq$profile-and$_.role-eq'candidate'-and$_.request_index-eq2})[0]
 $text.Add('| Recorded scope | Baseline ms [range] | Candidate ms [range] |')
 $text.Add('| --- | ---: | ---: |')
 foreach($metric in @('FolderView.ApplyCurrentSort.OrderItems','render.layout_items_us','render.begin_to_enddraw_us','render.present_us')){
  $bs=$b.stages.PSObject.Properties[$metric].Value
  $cs=$c.stages.PSObject.Properties[$metric].Value
  if($null-eq$bs-or$null-eq$cs){throw "Missing required attribution stage: $profile/$metric"}
  $text.Add("| $metric | $(Format-Ms $bs.median_us) [$(Format-Ms $bs.min_us) - $(Format-Ms $bs.max_us)] | $(Format-Ms $cs.median_us) [$(Format-Ms $cs.min_us) - $(Format-Ms $cs.max_us)] |")
 }
 $text.Add("| Request elapsed less counted writer spans | $(Format-Ms $b.remaining.median_us) [$(Format-Ms $b.remaining.min_us) - $(Format-Ms $b.remaining.max_us)] | $(Format-Ms $c.remaining.median_us) [$(Format-Ms $c.remaining.min_us) - $(Format-Ms $c.remaining.max_us)] |")
 $text.Add('')
 $text.Add('The arithmetic remainder is computed for each original request before grouping. It includes other application/test work, graphics waits and per-call microsecond quantization; it is not pure CPU time or an estimate of capture-disabled latency. Nested scopes overlap and must not be added together.')
 $text.Add('')
 $text.Add('Preparation remains separate from the request timer. The following retains every 5,000-item enumeration rank/sort-merge record, including counts, rather than selecting a fast initialization. These nested stage durations are not end-to-end folder-open latency.')
 $text.Add('')
 $text.Add('| Preparation stage | Baseline ms [range], count | Candidate ms [range], count |')
 $text.Add('| --- | ---: | ---: |')
 foreach($metric in @('FolderView.ExecuteEnumeration.NaturalSortRanks','FolderView.ExecuteEnumeration.SortMerge')){
  $b=@($data.preparation_statistics|Where-Object {$_.configuration-eq$profile-and$_.role-eq'baseline'-and$_.metric-eq$metric})[0].duration
  $c=@($data.preparation_statistics|Where-Object {$_.configuration-eq$profile-and$_.role-eq'candidate'-and$_.metric-eq$metric})[0].duration
  $text.Add("| $metric | $(Format-Ms $b.median_us) [$(Format-Ms $b.min_us) - $(Format-Ms $b.max_us)], $($b.count) | $(Format-Ms $c.median_us) [$(Format-Ms $c.min_us) - $(Format-Ms $c.max_us)], $($c.count) |")
 }
 $text.Add('')
 $text.Add('## Runs and observation')
 $text.Add('')
 $text.Add('| Role/block | Run ID | Receipt | Overlap polls | Max polling gap ms |')
 $text.Add('| --- | --- | --- | ---: | ---: |')
 foreach($phase in @($data.phases|Where-Object configuration -eq $profile)){
  $text.Add("| $($phase.phase) | $($phase.run_id) | $($phase.receipt_id) | $($phase.overlap_polls) | $($phase.maximum_gap_ms) |")
 }
 $text.Add('')
 $text.Add('All builds finish before timing. The shared observer requires a 30-second quiet gate and monitors the canonical runner window, nominally every 200 ms. It blocks compiler/build launchers, known DxUi test/benchmark executables, test drivers and RedPrism --self-test. Real already-running compiler and RedPrism self-test positive controls are retained; the generic probe label says synthetic, but neither control launched a test/compiler. Exact idle MSBuild identities require exited-parent and stable-CPU verification. Resumed CPU or a new tool blocks eligibility. Every rejected attempt is retained; acceptance never selects by latency.')
 $text.Add('')
 $text.Add('This is snapshot-based absence of detected compilation or known native-test overlap, not a lossless process-start audit, proof of an idle host, or physical scanout timing. Poll gaps, unknown shells, process exclusions and observer CPU remain explicit in each observation. Two authorized chats were asked to hold new work after current operations finish; a third chat was not messaged at the user''s direction. The observer waits for its compilation/tests to finish too; no independent process is interrupted.')
 $text.Add('')
 $text.Add('## Correctness, source and qualification')
 $text.Add('')
 $text.Add('Final test-enabled x64 Debug/Release builds and the final ten native cases per profile are separate from the superseded prototype. Native coverage includes complete JSONL schema/field order, controls and Unicode/malformed UTF-16, short-ASCII boundaries, uint64/HRESULT/counter semantics, immediate readback, concurrent rows, oversized scratch release, clear/reconfigure and same-path rotation. It also retains Windows ordering, current/selection, paint/no-op/failure/occlusion, rendering-error persistence and device-loss coverage. Source and tooling/format gates are recorded separately. No Full, ARM64 or ASan qualification is implied.')
 $text.Add('')
 $text.Add('Baseline source is a6ced5f41 with only identical test-only writer observation added to the old producer. Candidate C++ matches the user-facing checkout. Source preservation strips exactly the baseline diagnostic blocks and compares the remaining producer with a6ced5f41; all protected product/compiler/pin inputs match in all three checkouts. Final identities bind each compiled source snapshot, receipt and current runtime closure.')
 $text.Add('')
 $text.Add('A disk-full Z: volume caused preparation failures. The owned baseline and candidate checkouts were moved to E: and S:, then canonically restored and re-attested. Every final baseline/candidate run explicitly uses the same marked S:/RedSalamander.Perf fixture and capture root. Executable paths differ; loading/warm-up is outside measured requests. These fresh pairs must not be presented as a direct comparison with historical Z:-root absolute timings. Historical Z: sandbox findings remain untouched; the fresh S: audit has its own scope.')
 $text.Add('')
 $text.Add('The unattested first build, shared-dependency refusal, disk-full compilation/restore failures, stale generated-path recovery and superseded prototype receipts remain in attempt-ledger.json with digest-bound full logs. The first Debug native pass succeeded but its wrapper assumed the former Z: artifact location; exact S: artifacts were recovered without changing/rerunning that result. Final native runs use the corrected explicit root and final nonblocking observer.')
 $text.Add('')
 $text.Add('Completed attempts that overlap new compilation or exceed polling bounds are rejected even when native cases pass. Gates were also deliberately closed before any runner launched for bounded intervals granted to the already-authorized chat. Those interrupted gates have incomplete polling buffers and zero timing authority. Only first eligible completed blocks enter the comparison.')
 $text.Add('')
 $text.Add('One excluded Release-baseline attempt was a conservative classifier false positive: an otherwise read-only tooling inspection referenced test.ps1 in its command text. Its original flagged poll and full command remain retained. It is not described as actual compilation, and no eligibility override or latency-based reselection was applied; the next eligible complete block is used.')
 $text.Add('')
 $text.Add('After three accepted blocks, resumption preserves their exact fixed-order prefix instead of rerunning or reselecting samples. Fifteen guard tests plus actual verification cover source/receipt/profile/dependency identity, current compiled inputs/runtime closure under canonical profile locks, all required complete copied artifacts and the original polling rows/observer versions. Any absent ephemeral canonical source paths are explicitly listed in orchestrator-resumes.jsonl; the captured complete byte copies must still match the originally recorded hashes. No native code, timer, corpus or observer changes during this pause/resume. The original pre-resume orchestration source is retained too.')
 $text.Add('')
 $text.Add('## Evidence and reproduction')
 $text.Add('')
 $text.Add('artifact-bindings.json binds complete originals/extracts, byte counts, hashes and the exact selected metric families. All request samples are retained; full raw streams and all failed/rejected attempts remain under Z:/RedSalamander.Perf/evidence/debug-sort-product-cost-2026-10-20. Selected families retain every original row in order. Omitted per-item and other frame/preparation families remain intact externally and are included in analysis/metric counts. Receipt files are deduplicated by immutable ID. Archived orchestration scripts are evidence snapshots rather than supported repository commands.')
 $text.Add('')
 $text.Add('Complete process observations, including every gate and measurement row, use lossless .jsonl.gz storage to bound the packet size after long waits. Bindings include original and compressed hashes/lengths. Independent verification hashes the decompressed original bytes before checking all measurement polls; no observation row is filtered or dropped.')
 $text.Add('')
 $text.Add('```powershell')
 $text.Add('$env:REDSALAMANDER_TEST_ROOT=''S:\RedSalamander.Perf''')
 $text.Add("Tools/Run-AllTests.ps1 -Suite Commands -Configuration $profile -SkipBuild -CaseFilter folderView_perf_sort_toggle_stress -SelfTestRepeat 5")
 $text.Add('```')
 $text.Add('')
 $text.Add('[Independent verification](independent-verification.json)')
 [IO.File]::WriteAllText((Join-Path $packet 'README.md'),($text-join"`n")+"`n",[Text.UTF8Encoding]::new($false))
}
