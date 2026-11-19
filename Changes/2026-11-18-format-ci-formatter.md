- Made the shared root `format.ps1` RedSalamander's format command (step CV-6-05 of the `RedHome` convergence
  plan; register rows D-02, D-05, D-06 and E-02). It takes clang-format 22.1.3 from `Install-CiFormatter` in the
  pinned `.ci-engine` (cached under `.build/shared-format`), covers the files CI formats (the eight C and C++
  extensions under the `formatRoots` of `.ci/project.json`, minus its `formatExclude`, tracked or new and not
  ignored), formats by default and verifies with `-Check`. The repository's own pin and installer are removed:
  `Build/clang-format.json`, `Tools/Install-ClangFormat.ps1`, `Tools/Modules/Build/ClangFormat.psm1` and its
  `Tools/Tests/ClangFormat.Tests.ps1`. `format-all.ps1` stays for one release as a deprecated wrapper that runs
  `format.ps1`, maps `-WhatIf` to `format.ps1 -Check` and refuses a `-RepositoryRoot` other than its own checkout.
  The tool inventory, validation-impact rules, `Tools/README.md`, `Build_Toolchain.md`, the shared-helper and
  tooling-governance specs and the focused tests follow; the new fixture tests use the owned formatter probe and
  download nothing. 90 C and C++ files that the new scope found unformatted are reformatted, whitespace only: each
  is token-identical to its previous version once whitespace is removed, and no `#include` line moved. 25 further
  files that formatting would change beyond whitespace (include reordering, added namespace-closing comments or a
  split string literal) are left unformatted for a separate change, so `format.ps1 -Check` still lists them. No
  native build was run for this change.
