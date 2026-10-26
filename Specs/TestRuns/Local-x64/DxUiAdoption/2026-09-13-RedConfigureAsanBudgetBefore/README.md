# RedConfigure ASan budget diagnosis before correction

The same 512-token mass-preview fixture fails its 100 ms assertion in all five
isolated ASan processes on both master `0dd0bd6d` and migration `a20f6680`.
Master takes 215,101-220,010 us (median 217,737); migration takes 217,272-224,679 us
(median 219,852). These sequential blocks establish a pre-existing profile failure;
they do not establish paired performance non-regression. All raw outputs remain.

The fixture is identical apart from formatting; the workflow/model implementation
is unchanged. Master's ASan root build passes in 4m16s with one warning and no errors.
The original contract specifies normal Debug budgets. The correction preserves all
500/250/16/100 ms Debug/Release checks, retains every ASan functional assertion and
measurement, and labels ASan timings explicitly as diagnostic. Corrected runtime
validation remains pending. No product workflow or library pin is changed.

`comparison.json` lists every sample and exact executable hash. `retained-files.json`
records original-byte hashes and paths, including inert executed drivers. No hosted
CI ran. Earlier Full failures remain retained in the I19 qualification record.
