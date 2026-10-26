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
| Default | No regression established | within-noise-budget | 0 | 0 | 13 of 26 | 0 of 6 |
| MultilineGrid | Inconclusive | advice-required | 1 | 0 | 9 of 26 | 0 of 6 |
| MultilineGridDistinct | No regression established | within-noise-budget | 0 | 0 | 16 of 26 | 0 of 6 |

- **MultilineGrid**: 1 regressed metric has same-binary controls that drifted beyond their band on this runner, so the flag cannot be told from machine noise. Re-run the job; the flagged metrics are listed below, not dismissed.

### Flagged metrics

| Scenario | Phase | Metric | Baseline median | Candidate median | Change | p | Band | Same-binary controls | Outcome |
|---|---|---|---:|---:|---:|---:|---:|---|---|
| MultilineGrid | dirty | composeCpuP95Ms | 0.0166 | 0.0201 | +21.45% | 0.0281 | 5% | 22.4% max, beyond band | regressed, controls drifted |

<details><summary>Default: all 26 metrics</summary>

| Phase | Metric | Baseline median | Candidate median | Change | p | Band | Same-binary controls | Outcome |
|---|---|---|---:|---:|---:|---:|---|---|
| clean | fps | 353.1 | 305.6 | -13.46% | 0.1320 | 5% | 16.9% max, beyond band | within noise |
| clean | frameP50Ms | 2.8262 | 3.3406 | +18.20% | 0.0649 | 5% | 19.6% max, beyond band | within noise |
| clean | frameP95Ms | 2.9203 | 3.4353 | +17.64% | 0.0649 | 5% | 20.3% max, beyond band | within noise |
| clean | prepareP95Ms | 0.0005 | 0.0005 | 0.00% | 0.7056 | 5% | 50.0% max, beyond band | within noise |
| clean | composeCpuP95Ms | 0.0272 | 0.0288 | +5.89% | 0.1797 | 5% | 21.9% max, beyond band | within noise |
| clean | privateBytes | 23,414,784 | 23,367,680 | -0.20% | 0.4654 | 2% | 0.5% max, held | within noise |
| clean | privatePeakBytes | 23,414,784 | 23,408,640 | -0.03% | 0.7944 | 2% | 0.5% max, held | within noise |
| clean | workingSetBytes | 36,730,880 | 36,710,400 | -0.06% | 0.6991 | 2% | 3.2% max, beyond band | within noise |
| clean | workingSetPeakBytes | 36,749,312 | 36,728,832 | -0.06% | 0.6991 | 2% | 3.2% max, beyond band | within noise |
| clean | surfaceBytes | 3,686,400 | 3,686,400 | 0.00% | 1.0000 | exact | 0.0% max, held | within noise |
| clean | replacementPeakBytes | 3,686,400 | 3,686,400 | 0.00% | 1.0000 | exact | 0.0% max, held | within noise |
| clean | cppAllocations | 0 | 0 | 0.00% | 1.0000 | exact | 0.0% max, held | within noise |
| clean | composeAllocations | 0 | 0 | 0.00% | 1.0000 | exact | 0.0% max, held | within noise |
| dirty | fps | 167.0 | 165.2 | -1.06% | 0.5887 | 5% | 11.3% max, beyond band | within noise |
| dirty | frameP50Ms | 5.8501 | 5.9194 | +1.18% | 0.3095 | 5% | 3.2% max, held | within noise |
| dirty | frameP95Ms | 6.1779 | 6.8069 | +10.18% | 0.2403 | 5% | 16.8% max, beyond band | within noise |
| dirty | prepareP95Ms | 2.2413 | 2.2891 | +2.13% | 0.1212 | 5% | 26.5% max, beyond band | within noise |
| dirty | composeCpuP95Ms | 0.0132 | 0.0135 | +2.28% | 0.5671 | 5% | 22.2% max, beyond band | within noise |
| dirty | privateBytes | 23,654,400 | 23,672,832 | +0.08% | 0.8182 | 2% | 1.3% max, held | within noise |
| dirty | privatePeakBytes | 23,754,752 | 23,779,328 | +0.10% | 0.9740 | 2% | 1.3% max, held | within noise |
| dirty | workingSetBytes | 37,011,456 | 37,105,664 | +0.25% | 0.8528 | 2% | 3.6% max, beyond band | within noise |
| dirty | workingSetPeakBytes | 37,033,984 | 37,107,712 | +0.20% | 0.8528 | 2% | 3.6% max, beyond band | within noise |
| dirty | surfaceBytes | 3,686,400 | 3,686,400 | 0.00% | 1.0000 | exact | 0.0% max, held | within noise |
| dirty | replacementPeakBytes | 3,686,400 | 3,686,400 | 0.00% | 1.0000 | exact | 0.0% max, held | within noise |
| dirty | cppAllocations | 2,160 | 2,160 | 0.00% | 1.0000 | exact | 0.0% max, held | within noise |
| dirty | composeAllocations | 0 | 0 | 0.00% | 1.0000 | exact | 0.0% max, held | within noise |

</details>

<details><summary>MultilineGrid: all 26 metrics</summary>

| Phase | Metric | Baseline median | Candidate median | Change | p | Band | Same-binary controls | Outcome |
|---|---|---|---:|---:|---:|---:|---|---|
| clean | fps | 351.4 | 353.7 | +0.64% | 0.3095 | 5% | 36.6% max, beyond band | within noise |
| clean | frameP50Ms | 2.8071 | 2.805 | -0.07% | 0.4848 | 5% | 20.7% max, beyond band | within noise |
| clean | frameP95Ms | 2.975 | 2.9373 | -1.27% | 0.1320 | 5% | 80.5% max, beyond band | within noise |
| clean | prepareP95Ms | 0.0005 | 0.0004 | -11.11% | 0.1818 | 5% | 40.0% max, beyond band | within noise |
| clean | composeCpuP95Ms | 0.0261 | 0.0261 | +0.19% | 0.9372 | 5% | 39.8% max, beyond band | within noise |
| clean | privateBytes | 28,741,632 | 28,583,936 | -0.55% | 0.1970 | 2% | 3.4% max, beyond band | within noise |
| clean | privatePeakBytes | 28,774,400 | 28,616,704 | -0.55% | 0.1450 | 2% | 3.3% max, beyond band | within noise |
| clean | workingSetBytes | 41,656,320 | 41,609,216 | -0.11% | 0.0736 | 2% | 3.2% max, beyond band | within noise |
| clean | workingSetPeakBytes | 41,682,944 | 41,635,840 | -0.11% | 0.0931 | 2% | 3.1% max, beyond band | within noise |
| clean | surfaceBytes | 3,686,400 | 3,686,400 | 0.00% | 1.0000 | exact | 0.0% max, held | within noise |
| clean | replacementPeakBytes | 3,686,400 | 3,686,400 | 0.00% | 1.0000 | exact | 0.0% max, held | within noise |
| clean | cppAllocations | 0 | 0 | 0.00% | 1.0000 | exact | 0.0% max, held | within noise |
| clean | composeAllocations | 0 | 0 | 0.00% | 1.0000 | exact | 0.0% max, held | within noise |
| dirty | fps | 66.4 | 65.5 | -1.47% | 0.5887 | 5% | 11.5% max, beyond band | within noise |
| dirty | frameP50Ms | 14.5335 | 14.7785 | +1.69% | 0.1797 | 5% | 3.2% max, held | within noise |
| dirty | frameP95Ms | 16.0118 | 16.9058 | +5.58% | 0.4848 | 5% | 15.9% max, beyond band | within noise |
| dirty | prepareP95Ms | 12.0653 | 13.5028 | +11.91% | 0.1797 | 5% | 20.0% max, beyond band | within noise |
| dirty | composeCpuP95Ms | 0.0166 | 0.0201 | +21.45% | 0.0281 | 5% | 22.4% max, beyond band | regressed, controls drifted |
| dirty | privateBytes | 30,033,920 | 29,114,368 | -3.06% | 0.1320 | 2% | 12.1% max, beyond band | within noise |
| dirty | privatePeakBytes | 30,771,200 | 30,541,824 | -0.75% | 0.1320 | 2% | 2.2% max, beyond band | within noise |
| dirty | workingSetBytes | 42,692,608 | 41,734,144 | -2.25% | 0.2403 | 2% | 10.4% max, beyond band | within noise |
| dirty | workingSetPeakBytes | 43,096,064 | 42,901,504 | -0.45% | 0.3095 | 2% | 2.5% max, beyond band | within noise |
| dirty | surfaceBytes | 3,686,400 | 3,686,400 | 0.00% | 1.0000 | exact | 0.0% max, held | within noise |
| dirty | replacementPeakBytes | 3,686,400 | 3,686,400 | 0.00% | 1.0000 | exact | 0.0% max, held | within noise |
| dirty | cppAllocations | 1,080 | 1,080 | 0.00% | 1.0000 | exact | 0.0% max, held | within noise |
| dirty | composeAllocations | 0 | 0 | 0.00% | 1.0000 | exact | 0.0% max, held | within noise |

</details>

<details><summary>MultilineGridDistinct: all 26 metrics</summary>

| Phase | Metric | Baseline median | Candidate median | Change | p | Band | Same-binary controls | Outcome |
|---|---|---|---:|---:|---:|---:|---|---|
| clean | fps | 352.1 | 353.6 | +0.42% | 0.6991 | 5% | 8.1% max, beyond band | within noise |
| clean | frameP50Ms | 2.8059 | 2.8064 | +0.02% | 0.8182 | 5% | 3.3% max, held | within noise |
| clean | frameP95Ms | 2.9896 | 2.9525 | -1.24% | 0.2403 | 5% | 20.0% max, beyond band | within noise |
| clean | prepareP95Ms | 0.0005 | 0.0005 | 0.00% | 0.3939 | 5% | 25.0% max, beyond band | within noise |
| clean | composeCpuP95Ms | 0.0251 | 0.0273 | +8.57% | 0.3680 | 5% | 16.3% max, beyond band | within noise |
| clean | privateBytes | 25,296,896 | 25,366,528 | +0.28% | 0.7381 | 2% | 0.9% max, held | within noise |
| clean | privatePeakBytes | 25,333,760 | 25,430,016 | +0.38% | 0.8182 | 2% | 1.4% max, held | within noise |
| clean | workingSetBytes | 38,627,328 | 38,590,464 | -0.10% | 0.4848 | 2% | 0.6% max, held | within noise |
| clean | workingSetPeakBytes | 38,651,904 | 38,610,944 | -0.11% | 0.3680 | 2% | 0.6% max, held | within noise |
| clean | surfaceBytes | 3,686,400 | 3,686,400 | 0.00% | 1.0000 | exact | 0.0% max, held | within noise |
| clean | replacementPeakBytes | 3,686,400 | 3,686,400 | 0.00% | 1.0000 | exact | 0.0% max, held | within noise |
| clean | cppAllocations | 0 | 0 | 0.00% | 1.0000 | exact | 0.0% max, held | within noise |
| clean | composeAllocations | 0 | 0 | 0.00% | 1.0000 | exact | 0.0% max, held | within noise |
| dirty | fps | 160.1 | 158.3 | -1.12% | 0.9372 | 5% | 13.0% max, beyond band | within noise |
| dirty | frameP50Ms | 6.0613 | 6.0847 | +0.39% | 0.3939 | 5% | 1.1% max, held | within noise |
| dirty | frameP95Ms | 6.424 | 7.4209 | +15.52% | 0.3939 | 5% | 16.1% max, beyond band | within noise |
| dirty | prepareP95Ms | 2.9156 | 2.9377 | +0.76% | 0.8182 | 5% | 11.0% max, beyond band | within noise |
| dirty | composeCpuP95Ms | 0.0133 | 0.0134 | +0.75% | 0.7359 | 5% | 7.6% max, beyond band | within noise |
| dirty | privateBytes | 25,585,664 | 25,673,728 | +0.34% | 0.4848 | 2% | 2.4% max, beyond band | within noise |
| dirty | privatePeakBytes | 25,628,672 | 25,710,592 | +0.32% | 0.3074 | 2% | 2.2% max, beyond band | within noise |
| dirty | workingSetBytes | 39,184,384 | 39,122,944 | -0.16% | 0.3095 | 2% | 1.2% max, held | within noise |
| dirty | workingSetPeakBytes | 39,194,624 | 39,143,424 | -0.13% | 0.3095 | 2% | 1.2% max, held | within noise |
| dirty | surfaceBytes | 3,686,400 | 3,686,400 | 0.00% | 1.0000 | exact | 0.0% max, held | within noise |
| dirty | replacementPeakBytes | 3,686,400 | 3,686,400 | 0.00% | 1.0000 | exact | 0.0% max, held | within noise |
| dirty | cppAllocations | 1,080 | 1,080 | 0.00% | 1.0000 | exact | 0.0% max, held | within noise |
| dirty | composeAllocations | 0 | 0 | 0.00% | 1.0000 | exact | 0.0% max, held | within noise |

</details>

The retained receipts, comparisons and `summary.json` are in the `paired-benchmark-x64-Release` artifact. A metric regresses only when p < 0.05 and its median shift is beyond the band (a rise in an exact budget always regresses); `Specs/Core/Core_PerformanceAndResources.md` says what a verdict establishes and what a finding requires.
