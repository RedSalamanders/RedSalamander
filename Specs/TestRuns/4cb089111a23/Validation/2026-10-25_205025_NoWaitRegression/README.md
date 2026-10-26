# Owned warmup-lock query regression

Source e7f3b7188d9279e74fc1f9ac8d7f337d332445a3, x64 ASan Debug, receipt
94711f975ce6168ee9702f30290336c1209e507689c68b07d390ec177ee79cd3.
The solution built without warnings/errors. Ordinary sanitizer options and ordinary
NTFS seeding were used; no traversal override or timeout increase was selected.

The controlled INITIALIZING callback holds the repository volume mutex behind an
owned five-second release gate. Both query modes must return the exact live alpha.txt
result within 500 ms while the gate remains held. The authoritative route waited
5,011,577 us; the opportunistic route waited 5,014,767 us. Both returned the correct
file only after their gate timed out. The unchanged status snapshot passed. These
are two failed native query cases and a failed wrapper, never qualification.
The query fixture and its budgets remain unchanged across the production repair.

The same build's thirteen-case UI predecessor cluster passed every original case,
including conflict inspection return after B completes, retention, breadcrumb and
the 100 ms held-worker close test. Inspection focus is restored before binding;
keyboard authority is resolved from the current hosted control tree. The earlier
failed focus verdicts remain retained separately. A focused cluster does not close
three-profile Fresh Full, native platform, performance or manual acceptance gates.

Complete native and wrapper verdicts and the finished driver state are digest-bound
in manifest.json. The production no-wait lock repair still requires positive focused
checks, six builds and all final qualification. No failed Full receipt is promoted.
