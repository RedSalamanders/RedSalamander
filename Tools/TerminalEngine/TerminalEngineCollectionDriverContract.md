# Candidate-wrapper collection driver ABI

This ABI exists only for descriptor `build.kind=candidate-wrapper-v1`.
Provider-generic CMake, Zig, and Cargo builds never create an input, execute a
candidate driver, or accept a driver result.

The wrapper path is an exact changed path in the canonical candidate patch.
Its existence mechanically derives `maintainedBuildWrapper=true`; the provider
must compare that fact with the frozen assessment before building or scoring.
A wrapper cannot be renamed, moved into host code, or hidden behind a generic
kind to obtain `maintainedBuildWrapper=false`.

The provider writes at most 32,768 exact JCS bytes satisfying
`TerminalEngineCollectionDriverInput.schema.json` to a new file under the
fresh lane build root. The candidate identity binds the patch result tree and
collection-descriptor SHA. Only canonical absolute paths derived by the
provider are supplied. `sourceRoot`, notice/header/tool paths, the verified
candidate-neutral `hostHeaderIncludeRoot`, and resolved bindings are read-only
capabilities. `gate0.hostHeaderTreeSha256` binds that complete explicit tree;
the candidate cannot substitute ambient vcpkg headers. `buildRoot`,
`buildReceiptPath`, `sourceDependenciesPath`, and the separately supplied
result path are the only write authority.

Resolved bindings correspond one-for-one with descriptor `pathBindings`.
They carry no arbitrary command, environment value, tool path, or candidate
claim. The provider verifies each identity before launch. It freshly expands
the frozen host-header archive for each consuming phase and revalidates its
complete tree immediately before and after the isolated wrapper. A persisted
input mutation fails collection; the wrapper result cannot bless a changed
tree.

The wrapper runs once under the same network/process isolation and bounded
timeout as generic collection. It must synchronously reap every child and
close all handles before returning. Detached/background descendants,
post-return writes, pre-created provider verification output, reparse points,
and writes outside the lane root fail collection. Provider teardown always
terminates the isolated process tree before consuming a result.

The result path must be absent before launch. The wrapper creates at most 4,096
exact JCS bytes satisfying
`TerminalEngineCollectionDriverResult.schema.json`. A pass has
`failure=null`. A failure is exactly the frozen configure or build
phase/code pair. Stdout, stderr, exit status, timeout, and process-tree facts
come from provider isolation; candidate result text cannot replace them.

Artifact names in the input are discovery expectations only. The provider
independently discovers, hashes, inspects, and proves the adapter and complete
private PE closure. A result never asserts artifact identity or gate evidence.
