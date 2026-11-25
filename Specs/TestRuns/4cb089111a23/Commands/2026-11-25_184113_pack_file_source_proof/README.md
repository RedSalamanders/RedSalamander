# I66 Pack Delete after regular-file identity

The owned archive case packed one unchanged regular file through the 7z
writer and verified its cleanup. It also packed a different file into ZIP,
held Delete preparation, replaced the source at its original name and
verified that the task kept both the replacement and the moved original.
The same case passed the existing Unpack identity and partial-result checks.

- Test-enabled x64 Debug `3b7997fce` built with zero warnings and errors on
  receipt `af5274009b2d55cc2f09c5abddf1535e226af9f475f18ebe94d9cae9acf848d8`.
- Non-activating run `20261125T174113Z-85544-d590850a625548299c60c58c56680533`
  passed in 1,697 ms with zero failures and skips and archived normally.
  The focused repository verdict was `NOT_EVALUATED`.
- The ordinary stored ZIP Pack timing was 26,746 µs on Debug `5ac478f9b`
  and 17,938 µs on Debug `3b7997fce`. Both used the same 16-entry, 23-byte
  payload. Each timing comes from a single small archive.
- The initial compiler-free run `20261125T173500Z-72136-cbcfed09359142b7a74f514b6f53379d`
  failed one source check that matched the previous cleanup call. The
  corrected guard at `ea328f850` requires the central Delete submission
  and the exact cleanup parameter. Rerun
  `20261125T174237Z-89044-87e780002d2a4e1d9e06bf141ce56b42` passed all
  67 plan checks and 276 source and tooling checks.

Pack directory trees, archive workers and broader qualification remain open.
