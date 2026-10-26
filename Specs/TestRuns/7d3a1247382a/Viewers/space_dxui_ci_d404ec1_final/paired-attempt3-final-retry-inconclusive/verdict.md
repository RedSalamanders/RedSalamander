## Paired benchmark: Inconclusive

The run cannot say. A metric was flagged but its same-binary controls drifted beyond its band, or the set is too small to judge a timing. It does not pass: re-run the job for a new runner, or repeat the set on a quiet machine, and treat the flagged metrics as unresolved findings, not as noise.

| | |
|---|---|
| Baseline | `bea676a1f841` first parent of the merge commit (the base as the merge ref was made) |
| Candidate | `4c2aefd130ff` (this checkout) |
| Library inputs | changed, fingerprint `41423433DDF4` to `F826386B05B7` |
| Method | 3 repetitions of the interleaved pass A, B, B, A, so 6 runs per side; x64 Release; exact two-sided Mann-Whitney U test with the 5% timing and 2% memory investigation bands, exact budgets stay exact |
| Runner | runnervmfi6oq; a shared hosted VM, not a controlled quiet desktop |
| Event | pull request merge ref (refs/pull/69/merge) |

### Scenarios

| Scenario | Check | Set verdict | Regressed | Improved | Metrics held by controls | Controls stable |
|---|---|---|---:|---:|---:|---:|
| Default | Inconclusive | advice-required | 1 | 0 | 11 of 26 | 0 of 6 |
| MultilineGrid | No regression established | within-noise-budget | 0 | 0 | 10 of 26 | 0 of 6 |
| MultilineGridDistinct | Inconclusive | advice-required | 1 | 0 | 16 of 26 | 0 of 6 |

- **Default**: 1 regressed metric has same-binary controls that drifted beyond their band on this runner, so the flag cannot be told from machine noise. Re-run the job; the flagged metrics are listed below, not dismissed.

- **MultilineGridDistinct**: 1 regressed metric has same-binary controls that drifted beyond their band on this runner, so the flag cannot be told from machine noise. Re-run the job; the flagged metrics are listed below, not dismissed.

### Flagged metrics

| Scenario | Phase | Metric | Baseline median | Candidate median | Change | p | Band | Same-binary controls | Outcome |
|---|---|---|---:|---:|---:|---:|---:|---|---|
| Default | clean | fps | 277.1 | 244.5 | -11.76% | 0.0411 | 5% | 13.3% max, beyond band | regressed, controls drifted |
| MultilineGridDistinct | clean | prepareP95Ms | 0.0006 | 0.0007 | +16.67% | 0.0152 | 5% | 14.3% max, beyond band | regressed, controls drifted |

<details><summary>Default: all 26 metrics</summary>

| Phase | Metric | Baseline median | Candidate median | Change | p | Band | Same-binary controls | Outcome |
|---|---|---|---:|---:|---:|---:|---|---|
| clean | fps | 277.1 | 244.5 | -11.76% | 0.0411 | 5% | 13.3% max, beyond band | regressed, controls drifted |
| clean | frameP50Ms | 3.5816 | 4.1629 | +16.23% | 0.0649 | 5% | 9.9% max, beyond band | within noise |
| clean | frameP95Ms | 3.7486 | 4.3271 | +15.43% | 0.1797 | 5% | 36.1% max, beyond band | within noise |
| clean | prepareP95Ms | 0.0006 | 0.0008 | +36.36% | 0.1429 | 5% | 40.0% max, beyond band | within noise |
| clean | composeCpuP95Ms | 0.0223 | 0.0312 | +39.91% | 0.4848 | 5% | 36.9% max, beyond band | within noise |
| clean | privateBytes | 23,420,928 | 23,494,656 | +0.31% | 0.6991 | 2% | 1.4% max, held | within noise |
| clean | privatePeakBytes | 23,441,408 | 23,523,328 | +0.35% | 0.4004 | 2% | 1.3% max, held | within noise |
| clean | workingSetBytes | 36,700,160 | 36,886,528 | +0.51% | 0.8182 | 2% | 3.5% max, beyond band | within noise |
| clean | workingSetPeakBytes | 36,722,688 | 36,909,056 | +0.51% | 0.8182 | 2% | 3.5% max, beyond band | within noise |
| clean | surfaceBytes | 3,686,400 | 3,686,400 | 0.00% | 1.0000 | exact | 0.0% max, held | within noise |
| clean | replacementPeakBytes | 3,686,400 | 3,686,400 | 0.00% | 1.0000 | exact | 0.0% max, held | within noise |
| clean | cppAllocations | 0 | 0 | 0.00% | 1.0000 | exact | 0.0% max, held | within noise |
| clean | composeAllocations | 0 | 0 | 0.00% | 1.0000 | exact | 0.0% max, held | within noise |
| dirty | fps | 144.8 | 150.2 | +3.76% | 0.9372 | 5% | 39.9% max, beyond band | within noise |
| dirty | frameP50Ms | 6.498 | 6.385 | -1.74% | 0.4848 | 5% | 23.4% max, beyond band | within noise |
| dirty | frameP95Ms | 6.9714 | 8.1354 | +16.70% | 0.4848 | 5% | 37.5% max, beyond band | within noise |
| dirty | prepareP95Ms | 2.0527 | 2.203 | +7.32% | 0.9372 | 5% | 23.1% max, beyond band | within noise |
| dirty | composeCpuP95Ms | 0.0172 | 0.0167 | -2.92% | 0.3723 | 5% | 10.8% max, beyond band | within noise |
| dirty | privateBytes | 23,728,128 | 23,867,392 | +0.59% | 0.4610 | 2% | 1.7% max, held | within noise |
| dirty | privatePeakBytes | 23,810,048 | 23,879,680 | +0.29% | 0.8182 | 2% | 2.1% max, beyond band | within noise |
| dirty | workingSetBytes | 37,056,512 | 37,144,576 | +0.24% | 0.9372 | 2% | 3.7% max, beyond band | within noise |
| dirty | workingSetPeakBytes | 37,072,896 | 37,146,624 | +0.20% | 0.9372 | 2% | 3.7% max, beyond band | within noise |
| dirty | surfaceBytes | 3,686,400 | 3,686,400 | 0.00% | 1.0000 | exact | 0.0% max, held | within noise |
| dirty | replacementPeakBytes | 3,686,400 | 3,686,400 | 0.00% | 1.0000 | exact | 0.0% max, held | within noise |
| dirty | cppAllocations | 2,160 | 2,160 | 0.00% | 1.0000 | exact | 0.0% max, held | within noise |
| dirty | composeAllocations | 0 | 0 | 0.00% | 1.0000 | exact | 0.0% max, held | within noise |

</details>

<details><summary>MultilineGrid: all 26 metrics</summary>

| Phase | Metric | Baseline median | Candidate median | Change | p | Band | Same-binary controls | Outcome |
|---|---|---|---:|---:|---:|---:|---|---|
| clean | fps | 274.0 | 272.9 | -0.40% | 0.5887 | 5% | 1.7% max, held | within noise |
| clean | frameP50Ms | 3.6204 | 3.6175 | -0.08% | 0.9740 | 5% | 0.7% max, held | within noise |
| clean | frameP95Ms | 3.8262 | 3.8708 | +1.17% | 0.5887 | 5% | 6.4% max, beyond band | within noise |
| clean | prepareP95Ms | 0.0007 | 0.0007 | 0.00% | 0.4545 | 5% | 16.7% max, beyond band | within noise |
| clean | composeCpuP95Ms | 0.025 | 0.0249 | -0.60% | 0.7316 | 5% | 16.3% max, beyond band | within noise |
| clean | privateBytes | 28,782,592 | 28,672,000 | -0.38% | 0.6212 | 2% | 3.1% max, beyond band | within noise |
| clean | privatePeakBytes | 28,819,456 | 28,715,008 | -0.36% | 0.6667 | 2% | 3.0% max, beyond band | within noise |
| clean | workingSetBytes | 41,768,960 | 41,699,328 | -0.17% | 0.3095 | 2% | 3.0% max, beyond band | within noise |
| clean | workingSetPeakBytes | 41,799,680 | 41,725,952 | -0.18% | 0.3095 | 2% | 3.0% max, beyond band | within noise |
| clean | surfaceBytes | 3,686,400 | 3,686,400 | 0.00% | 1.0000 | exact | 0.0% max, held | within noise |
| clean | replacementPeakBytes | 3,686,400 | 3,686,400 | 0.00% | 1.0000 | exact | 0.0% max, held | within noise |
| clean | cppAllocations | 0 | 0 | 0.00% | 1.0000 | exact | 0.0% max, held | within noise |
| clean | composeAllocations | 0 | 0 | 0.00% | 1.0000 | exact | 0.0% max, held | within noise |
| dirty | fps | 58.5 | 57.9 | -0.93% | 0.8182 | 5% | 18.2% max, beyond band | within noise |
| dirty | frameP50Ms | 16.3863 | 16.6832 | +1.81% | 0.2403 | 5% | 12.8% max, beyond band | within noise |
| dirty | frameP95Ms | 18.1233 | 19.6228 | +8.27% | 0.0931 | 5% | 13.1% max, beyond band | within noise |
| dirty | prepareP95Ms | 13.4972 | 14.488 | +7.34% | 0.1797 | 5% | 10.4% max, beyond band | within noise |
| dirty | composeCpuP95Ms | 0.0182 | 0.0196 | +7.71% | 0.2338 | 5% | 25.3% max, beyond band | within noise |
| dirty | privateBytes | 29,380,608 | 29,550,592 | +0.58% | 1.0000 | 2% | 15.7% max, beyond band | within noise |
| dirty | privatePeakBytes | 30,633,984 | 30,568,448 | -0.21% | 0.8550 | 2% | 2.4% max, beyond band | within noise |
| dirty | workingSetBytes | 42,326,016 | 42,092,544 | -0.55% | 0.8182 | 2% | 14.0% max, beyond band | within noise |
| dirty | workingSetPeakBytes | 43,057,152 | 42,962,944 | -0.22% | 0.4848 | 2% | 3.4% max, beyond band | within noise |
| dirty | surfaceBytes | 3,686,400 | 3,686,400 | 0.00% | 1.0000 | exact | 0.0% max, held | within noise |
| dirty | replacementPeakBytes | 3,686,400 | 3,686,400 | 0.00% | 1.0000 | exact | 0.0% max, held | within noise |
| dirty | cppAllocations | 1,080 | 1,080 | 0.00% | 1.0000 | exact | 0.0% max, held | within noise |
| dirty | composeAllocations | 0 | 0 | 0.00% | 1.0000 | exact | 0.0% max, held | within noise |

</details>

<details><summary>MultilineGridDistinct: all 26 metrics</summary>

| Phase | Metric | Baseline median | Candidate median | Change | p | Band | Same-binary controls | Outcome |
|---|---|---|---:|---:|---:|---:|---|---|
| clean | fps | 272.4 | 269.1 | -1.22% | 0.4848 | 5% | 10.6% max, beyond band | within noise |
| clean | frameP50Ms | 3.6208 | 3.6337 | +0.36% | 0.8182 | 5% | 2.7% max, held | within noise |
| clean | frameP95Ms | 3.7945 | 3.8934 | +2.61% | 0.3095 | 5% | 26.2% max, beyond band | within noise |
| clean | prepareP95Ms | 0.0006 | 0.0007 | +16.67% | 0.0152 | 5% | 14.3% max, beyond band | regressed, controls drifted |
| clean | composeCpuP95Ms | 0.0244 | 0.0262 | +7.17% | 0.2814 | 5% | 21.1% max, beyond band | within noise |
| clean | privateBytes | 25,374,720 | 25,354,240 | -0.08% | 0.8182 | 2% | 1.2% max, held | within noise |
| clean | privatePeakBytes | 25,446,400 | 25,391,104 | -0.22% | 0.6991 | 2% | 1.2% max, held | within noise |
| clean | workingSetBytes | 38,760,448 | 38,594,560 | -0.43% | 0.0649 | 2% | 0.7% max, held | within noise |
| clean | workingSetPeakBytes | 38,778,880 | 38,608,896 | -0.44% | 0.0649 | 2% | 0.7% max, held | within noise |
| clean | surfaceBytes | 3,686,400 | 3,686,400 | 0.00% | 1.0000 | exact | 0.0% max, held | within noise |
| clean | replacementPeakBytes | 3,686,400 | 3,686,400 | 0.00% | 1.0000 | exact | 0.0% max, held | within noise |
| clean | cppAllocations | 0 | 0 | 0.00% | 1.0000 | exact | 0.0% max, held | within noise |
| clean | composeAllocations | 0 | 0 | 0.00% | 1.0000 | exact | 0.0% max, held | within noise |
| dirty | fps | 139.0 | 142.5 | +2.53% | 0.2403 | 5% | 10.6% max, beyond band | within noise |
| dirty | frameP50Ms | 6.823 | 6.802 | -0.31% | 0.9372 | 5% | 1.7% max, held | within noise |
| dirty | frameP95Ms | 7.3398 | 8.3831 | +14.21% | 0.2403 | 5% | 14.1% max, beyond band | within noise |
| dirty | prepareP95Ms | 2.9613 | 2.9279 | -1.13% | 1.0000 | 5% | 9.2% max, beyond band | within noise |
| dirty | composeCpuP95Ms | 0.0157 | 0.0157 | 0.00% | 0.9416 | 5% | 9.5% max, beyond band | within noise |
| dirty | privateBytes | 25,796,608 | 25,714,688 | -0.32% | 0.2597 | 2% | 2.5% max, beyond band | within noise |
| dirty | privatePeakBytes | 25,923,584 | 25,767,936 | -0.60% | 0.2857 | 2% | 2.4% max, beyond band | within noise |
| dirty | workingSetBytes | 39,301,120 | 39,096,320 | -0.52% | 0.0411 | 2% | 1.5% max, held | within noise |
| dirty | workingSetPeakBytes | 39,337,984 | 39,114,752 | -0.57% | 0.0260 | 2% | 1.5% max, held | within noise |
| dirty | surfaceBytes | 3,686,400 | 3,686,400 | 0.00% | 1.0000 | exact | 0.0% max, held | within noise |
| dirty | replacementPeakBytes | 3,686,400 | 3,686,400 | 0.00% | 1.0000 | exact | 0.0% max, held | within noise |
| dirty | cppAllocations | 1,080 | 1,080 | 0.00% | 1.0000 | exact | 0.0% max, held | within noise |
| dirty | composeAllocations | 0 | 0 | 0.00% | 1.0000 | exact | 0.0% max, held | within noise |

</details>

The retained receipts, comparisons and `summary.json` are in the `paired-benchmark-x64-Release` artifact. A metric regresses only when p < 0.05 and its median shift is beyond the band (a rise in an exact budget always regresses); `Specs/Core/Core_PerformanceAndResources.md` says what a verdict establishes and what a finding requires.
