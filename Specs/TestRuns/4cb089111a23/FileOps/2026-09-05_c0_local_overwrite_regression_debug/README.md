# C0 direct Local overwrite regression

Run `c0-local-overwrite-regression-20260905`, Debug,
`Floodgate_LocalCopyOverwriteIsStaged`: 3/3 passed, no skips (one behavior plus setup/cleanup).
Build receipt `f8134788c8e746e61a4785969ee51f835f51c03dbeb39deb69295fe76b5786b1`;
zero warnings/errors. Base `58f22d0e` plus working changes is runner-bound.

The real Local copy API preserves the original during staged failure and verifies
replacement bytes after the corrected post-publication attributes fault. Results,
trace, aggregate and the complete bounded metric stream are unchanged raw artifacts.
This is correctness evidence, not a throughput claim or plan closeout.

The disk audit reports 24 retained earlier run directories outside its allow-list;
they were not deleted. This does not establish a clean test-root audit.
