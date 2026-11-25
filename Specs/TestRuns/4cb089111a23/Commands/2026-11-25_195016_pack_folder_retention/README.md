# I66 Pack keeps folder selections without a cleanup receipt

The owned archive case packed a regular file, a populated folder and an
empty folder with Delete after selected. The ZIP held the same complete
entry set as ordinary Pack. All originals stayed, and Pack asked no
permanent-delete question and queued no Delete. The omitted-link guard
and exact regular-file Pack and Unpack cleanup checks also passed.

- Test-enabled x64 Debug `78d2a5c7236914661d861f3f89dfd647e52ccef3` built
  with zero warnings and errors on receipt
  `e7f7ba5b552cf3b6fe0bcda7638d18d96c123d0a5ce4405ee5764ae4eaad0952`.
  The first build was interrupted at its slow owned Zig download. The
  retry verified all 11 cached dependencies against their pinned hashes.
- Non-activating ZIP case `20261125T185016Z-37604-8000aef41f024e65870c2d6df59b7de3`
  passed in 8,425 ms with zero failures and skips and archived successfully.
  The focused repository verdict was `NOT_EVALUATED`.
- Compiler-free run `20261125T184206Z-91560-9b6ffca17a80419a95cc7907efb2f4bf`
  passed 67 runner-plan and 276 source and tooling checks.
- The ordinary stored ZIP phase recorded 20,274 µs over 16 entries and
  23 bytes. This is one small-archive observation.

Exact descendant cleanup, worker migration, and broader qualification
remain open.
