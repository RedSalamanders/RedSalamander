# Per-item cancellation fixture baseline

This is failed baseline evidence from source
`56232e438b32953ddf9829529b87e7f3d59f1e5e`, x64 Release. The same executable,
build receipt, Full environment and approved E sandbox were used for the replay.
The prior Debug Fresh Full passed; these Release failures remain failed and do
not establish repository qualification. Candidate qualification is pending.

The broad run passed 17 other groups but FileOps stopped at its cancellation
fixture: 16 cases passed, one failed, and 305 downstream cases were blocked.
The unchanged native replay passed 17 repetitions before repetition 18 failed;
12 later repetitions were blocked. The copied full result files preserve every
case verdict and reason. Blocked skips are not prerequisite coverage.

The broad failure reported a Copy root as Published/Retained/Canceled. The replay
reported a Move root as Published/Removed/Canceled. All never-dispatched roots in
both observations remained NotAttempted/Retained/Canceled. The fixture assumed
visible destination bytes plus 300 ms proved completion; it had no terminal
witness before answering Cancel. The owning file-operation contract permits
published work to retain a later cancellation, independently of its storage axes.

`baseline-observations.json` retains the eight per-root observations and rejecting
assertion in their original order for each failure, with original trace SHA-256,
byte and row counts and line numbers. Complete raw traces remain in the named
local completed runs. `manifest.json` binds all three curated JSON payloads.
This is a synchronization repair baseline, not a product performance comparison
or proof that every cancellation path is correct.
