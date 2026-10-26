# Preceding File Operations capture regressions

Both exact native scenarios passed, each with setup and cleanup (three passed,
zero failed/skipped). Their source snapshot and verified App receipt ID remain
in the driver results. The later changes concern only the standalone playback
fixture and documentation. The App and shared screenshot-helper native bytes
remain identical, as the equivalence record verifies.

Complete raw runs have independent hash-verified private ZIP readbacks. These
results retain their original source identity and are not final-build or Full
qualification. The original driver expected one result and incorrectly marked
the first three-pass setup/scenario/cleanup result as a driver failure; native
results were successful throughout. That private driver check was corrected.
