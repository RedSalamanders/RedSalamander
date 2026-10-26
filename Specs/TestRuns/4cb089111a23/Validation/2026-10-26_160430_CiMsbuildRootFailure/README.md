# Custom-root MSBuild discovery failure and repair

At public head `1edb7f3c6aa6879cc86f13f3e409e0a2977a1ac3`, both native
CI jobs successfully provisioned the required toolchain and completed dependency
installation. Both then failed before C++ compilation because build.ps1 required
the literal Microsoft Visual Studio parent-directory name. The isolated hosted
installation correctly lives below RUNNER_TEMP. Complete job logs are retained;
these failed CI jobs are not promoted.

Discovery now verifies actual VS18+ version metadata, invariant MSBuild.exe binary
identity and the standard MSBuild/Current/Bin host layout. A custom installation
root is accepted; wrong layouts, other executables and missing files are rejected.
The required MSVC14.52 policy is unchanged. A real relocated MSBuild executable
and negative controls exercise the actual discovery function.

The first local tooling pass had one existing optional version-database skip
because the fresh checkout lacked the pinned vcpkg tool fixture. That record is
retained. After restoring the exact d53f25d tool checkout, all 95 focused tests
passed with zero failures/skips; canonical source identity remained unchanged
through that run. Read-only x64/ARM64 toolchain verification also passed. These
are dirty-worktree focused receipts, not passing results for the predecessor Git
commit. Updated public CI remains required. C++ sources and compiler/linker policy
are unchanged; Full runtime reruns remain omitted under the accepted scope.
