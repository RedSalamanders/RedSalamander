# Residual failure diagnostics and retained seed baselines

These receipts belong to source `fac05d5b70b9a5933b12093c1b2337f76aee8cdc`.
Its full x64 ASan rebuild and x64 Release build passed with zero warnings/errors.
They do not replace the required final six builds or Fresh Full profiles.

The six-case ViewerSpace predecessor cluster passed 18 checks across three
repetitions, including the strengthened current-presented-generation condition
and every original draw/LOD assertion. All Close checks passed in this focused
process; the earlier 133173-us Full failure remains open. Fine samples put
31–36 ms in the final brush-release span, while explicit D3D device reset and
parent nonclient cleanup were near zero. Windows child teardown remains inside
the whole-HWND endpoint. No asynchronous device-retirement repair is qualified.

The normal native-NTFS stale-root diagnostic recorded one local-index pass and
one service warmup timeout. The completed local seed collected 4,169,294 MFT
entries and 413,378 parent lists: 89,327,289 us enumeration, 1,314 us tracked-root
projection, 18,169,793 us temporary retirement, 107,498,396 us total. The service
case kept its original 60-second watchdog and failed. No traversal override,
timeout multiplier or sanitizer-quarantine adjustment was used.

Three fresh-process local-index baselines per profile retained normal MFT
enumeration, exact alpha/beta results, stale-store refresh and persistence guards.
Native MFT totals and 250-ms sampled private-memory peaks were:

| Profile | Native totals (us) | Sampled peak private bytes |
| --- | --- | --- |
| ASan Debug | 95699466, 98027650, 95204072 | 3199959040, 3199954944, 3199733760 |
| Release | 8881990, 8791645, 8877969 | 844038144, 847822848, 831844352 |

The actual enumeration counts grew from 4,169,462 to 4,169,695 as owned run
artifacts were added. These are the same machine, volume and owned scenario;
they are not a frozen image of every volume record. Sampling peaks are observed
peaks, not a guarantee that every transient allocation peak was captured.
The sampler records only exact-path descendants of its launching probe and
never samples or terminates the separately launched Z:/src Release app.

`manifest.json` binds complete native verdicts, baseline metrics, sampled memory,
source/artifact receipts and exact launcher/sampler sources. The 10-MiB Commands
stream is represented by ordered selected close rows with source digest and
line numbers; it is explicitly not the complete original stream. The earlier
b2cb close extract retains its distinct source identity.

The first external launch produced no native verdict: its expected path doubled
the native output suffix, and observer streams were placed inside the native
harness's rotating folder. Its raw partial state/error remain diagnostic only.
The corrected launcher places observer files beside the native output and
reuses the canonical runner's desktop mutex/execution-state function by AST,
plus the existing contained-process launcher. It does not create Full evidence.
The initially rejected non-rebuild attempt and required successful ASan rebuild
are retained separately. No failed or incomplete attempt is promoted.

These baselines precede a proposed scoped-arena optimization of temporary MFT
storage. Candidate correctness, paired ASan/Release timing/memory, the original
service watchdog, the remaining Close repair, final Full qualification and the
public migration all remain pending.
