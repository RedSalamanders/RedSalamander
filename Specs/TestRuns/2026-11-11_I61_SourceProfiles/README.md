# I61 current-source build profiles

Historical record: the source below was current when these builds ran. These
receipts do not qualify the later merged `main` or its current toolchain inputs.

[Qualification](qualification.json) records all six test-enabled Source builds of c45f3deba2b9edf555bc294b8775d9f56fdb2565, with the private compiler metadata factory at both direct and streaming contained launch boundaries. Debug, Release and ASan Debug built on x64 and ARM64 with zero warnings and errors; each exact receipt was byte-verified for artifact use immediately after building. Raw logs and all six receipts are preserved under the marked local evidence root.

This is build qualification only. It does not establish runtime, native ARM64 runtime, aggregate SDK Full, hosted or controlled performance acceptance. The earlier local SDK Full remains FAILED (2815 passed,3 failed,56 skipped); its fragment and two ViewerPE failure modes are separately tracked. Focused retries and evidence-reader controls do not supersede those failures.
