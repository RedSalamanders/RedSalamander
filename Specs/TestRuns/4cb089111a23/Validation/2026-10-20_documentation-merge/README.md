# Documentation and screenshot merge validation

This is fresh focused x64 Debug validation after applying the repository-pinned
clang-format 22.1.3 before the documentation commit. All 23 changed C++ files pass
the same whole-file formatting check as CI; thirteen corrections are verified
byte-for-byte as the formatter's output from preserved pre-format source.

The App and ViewerPETests builds complete with zero warnings/errors. The exact
App gallery passes all 45 scenes, and the standalone viewer, Space lifecycle and
real VLC playback/pause cases pass with no skips. Both focused File Operations
scenarios pass, each including setup and cleanup (3/0/0). Every native run uses
the existing no-activation path with no global keyboard/mouse input. Playback
retains the software-decoding/wingdi limits and both pixel-probe negative controls.

Both component receipts bind source snapshot `ffb7a8218c28a7064be534d46116ab944c766b7141e49be32bb4254a4c37608b`.
All changed source bytes and output digests matched those validation origins
at native closeout. Later curation and its evidence-only `.gitattributes` rule
are outside that exact source identity; native source and output digests remain
unchanged before commit. The 432 relevant Pester checks pass in two fresh disjoint
batches after formatting. Native raw runs have independent private ZIP readbacks.

All 98 public images remain byte-identical to the October 19 privacy/OCR review;
all 76 native public frames still match their recorded archived native origins.
The original published screenshot receipts remain unchanged. Formatting validation
does not rewrite their origin or restamp earlier evidence as this build.

PrePush explanation cannot delegate the changed reviewed tooling closure to its
existing PR workflow digest. It lists five foreground-input entries as pending.
No Full, Release, ARM64, foreground-input or clean-commit CI pass is claimed here.
The enabled master push workflow is expected to validate its committed candidate;
its future result is pending, never counted as a local pass.

The commit includes code, tests, normative specifications, user/developer guides,
synthetic screenshots, provenance and useful rejected-attempt/validation evidence.
Private raw evidence and the unrelated theme branch/worktree are preserved.

Six remaining exact owned runs were hash-archived with independent ZIP readback
and removed through the public sandbox cleaner. The final sandbox disk audit
reports zero issues. No unrelated worktree, branch, process or private evidence
was removed. Documentation checks resolve all 762 local references across 47
pages and cover 25 settings sections and 16 plugin families. The spec/tooling
inventories report no errors or warnings.

