# Hosted CI bootstrapper failure and GUI-process regression

PR #142 head `6dbeba14755fd59c39c3e23292651e075c5c7858` passed formatting
and migration guards, but both native Debug jobs in CI run 40344755046 failed
before compilation. Their complete job logs retain the unset-LASTEXITCODE error.
The GUI bootstrapper returned asynchronously; the script read a shell exit
variable that had not been set. These failed CI jobs are not promoted.

The action now waits through the existing contained streaming-process helper,
reads its actual exit result, accepts only 0 or 3010, then verifies the required
tools. Its signature, hosted-runner, compiler-family and local-install refusal
checks remain intact. The new regression compiles a real Windows GUI fixture,
then executes the production hosted-provisioning branch using fixture download,
signature and tool-discovery seams. It proves delayed completion, exact Unicode
and spaced arguments, 0/3010 acceptance and failure 17 rejection with no software
installation or desktop window. The fixture PE subsystem is explicitly checked.

All 40 focused tooling cases passed without failures/skips. Both x64/ARM64
read-only toolchain checks passed. The local record binds the dirty worktree's
canonical source snapshot and verifies that identity remained unchanged during
the tests; it is not a passing receipt for the predecessor Git commit. Updated
public CI remains required. The qualified C++ source and its 51 native checks are
unchanged; Full runtime reruns remain omitted under the user's accepted scope.
