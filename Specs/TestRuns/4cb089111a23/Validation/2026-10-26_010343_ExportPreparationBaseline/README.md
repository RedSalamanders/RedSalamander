# ViewerSpace export preparation baseline

All runs precede the production preparation change and bind source
bf2ad18715abc6a406ae8b0ac302703d06b52617, exact native executable/plugin
digests, approved E: fixtures and default ASan settings. The x64 ASan Debug and
Release full-solution builds passed. These are focused baselines, not repository
qualification, and they do not clear either earlier failed/incomplete Full.

The two unchanged registered analysis cases retain 2002 real retained rows, their
long synthetic root, JSON/CSV writers, failed publication preservation, actual
host detach, worker gate, Close/unload vote, retirement and no-owned-activation
assertions. Sixteen traced and sixteen untraced ASan executions passed; four
traced Release executions passed. All 36 complete native outputs are retained.
The original five-second PumpUntil wait and its final message drain are unchanged.
Its measured ASan spans were 4.520675-4.862488 seconds with capture and
4.483983-5.013981 seconds without capture, including a passing final-drain
observation just after five seconds. Release spans were 0.243883-0.288795 seconds.
These passes do not prove the precise cause of the earlier uninstrumented failure.

The existing export preparation metrics show all four 2002-row preparations in
each traced execution. In the first ASan case, the held export takes 146 timer
turns, 195570 us of measured active preparation and a 4713 ms metric wall span;
the median interval between turns is 31 ms. Native gate time uses steady_clock;
metric wall spans use recorded UTC and are labeled separately. Tracing has its
own observer cost. The source constructs display text/attributes merely to test
timestamp validity and copies owned UTF-8 results; those avoidable costs are the
candidate's measured optimization target, without changing the 1.2 ms turn target.

The archive contains every native output, exact build receipts, launch/analysis
scripts as text, all per-run preparation summaries and original trace digests.
One complete representative trace per profile is retained to keep the archive
bounded; remaining complete traces remain at their recorded approved run paths.
No passing repetition, baseline or summary is promoted to Full qualification.
