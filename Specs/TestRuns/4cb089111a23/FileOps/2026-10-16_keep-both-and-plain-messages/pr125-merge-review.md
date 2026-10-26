# PR #125 archive review, 2026-10-17

The original two local archives were omitted when PR #125 was created. This change
promotes all 104 original files without modifying their bytes, including failed
attempts. Their README files describe the original runs and their limitations.

The final safety and conflict runs both bind receipt
`8f19e1856a6488bda35098c0c3a606ffa9db595a379cc37062046c13fd3d7c3d`
and workspace snapshot
`dc4a5a6d62557c40a9fc3faf68e2f470655082b2703d80b6d5f898107f73c906`.
Their result files contain 8 and 33 passed case executions, respectively, with
no failures or skips. The saved resource/harness result contains 214 passes.

The source comparison reconstructs Git tree
`df615af551fa41647d936209e6b97733fd296fd0` from base
`a48e7d7880cb37ea6898d462ae188a6213a5f3f8` and the archived source diff.
Only the comparison copy of the diff has its CRLF transport endings normalized;
the archived diff remains intact. All 23 recorded changed-file hashes match
that reconstructed source after applying the original checkout line endings.
Compared with PR #125's reviewed head
`fe8bd298986aa701a45069b0fb88445370cf7043`, the only differences are the
Backspace implementation, Commands navigation test, and FolderView specification
already merged in `ee433475521e8031501183450644d8a07bdd8769`.
The File Operations implementation, fixtures, resources, specifications, and
the remaining dependency tree are identical. The accompanying JSON records the
comparison and checks the closing metric streams against their retained raw inputs.

This is historical focused x64 Debug evidence for the recorded artifact. It is
not an exact receipt reuse, validation-entry promotion, new runtime test, or
qualification of the combined tree. It supplies no Release, ARM64, ASan, Full,
PR, live NAS, matched performance baseline, or speedup result. The original
Release-process blocker is historical. GitHub's current PR checks did not start
because of account billing/spending limits; they did not provide a build result.
