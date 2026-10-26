# Final focused Debug: Skip counting presentation

`FileOps_SkipCountingPresentation` **passed**. The runner recorded 3 passed, 0 failed,
0 skipped (Setup, the selected case, Cleanup), exit code 0.

- Run: `20261015T144632Z-48204-c817856820704371b0f22ba2c8ab610b`.
- [Test-enabled x64 Debug receipt](build-receipt.json): `047f98907061f915d00f1f490b967e89fab62bf15c3f8aa6169c531d0240b5af`, shared with the final [provider/count checks](../2026-10-15_164631-FileOps_PermanentDeleteCompletionProgress/README.md) and [repeated-choice checks](../2026-10-15_164725-FileOps_RepeatDecisionContracts/README.md).
- **36 layout configurations**: 9 scenarios × fr-FR/en-US × 480/760 DIP. The **18 harness captures** use 480 DIP at native **144 DPI**, without activating the test window.

The scenarios cover exact/file-only/bytes-only Copy summaries, known/unknown Delete counts,
direct DiskFull choices, a full choice cache, read-only-folder deletion consent, and DiskFull
without a qualified destination. Checks cover the one painted work summary, full accessible
text, clipping/action-row boundaries, direct actions and omitted inapplicable controls.
The 10 work-summary captures also pass color probes.

The [manifest](captures/scenarios.tsv) records scenario, language, width, DPI and compile time.
Representative actual application captures: [Copy](captures/fr-FR-copy-counts.png),
[Delete](captures/fr-FR-delete-counts.png), [Disk full](captures/en-US-disk-full.png),
[unknown Delete counts](captures/en-US-delete-counts-unavailable.png), and
[read-only folder consent](captures/fr-FR-read-only-folder-delete.png).

These deterministic snapshots validate production-renderer presentation; live provider counters
are covered separately by the linked Delete/bridge case. The [raw aggregate](run-all-tests-results.json)
retains `is_clean:false` and all 3 legacy sandbox-directory findings. Raw captures, traces,
results and statuses remain unchanged. This focused Debug pass is not a Full-repository, Release,
ARM64, remote-provider or foreground-suite verdict, and makes no throughput-improvement claim.
