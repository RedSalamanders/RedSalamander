# Export diagnostic build failure

The x64 ASan Debug rebuild of 3ff97a52c8e171aba121b0598e07fdb1f9268154
failed ViewerPETests linking after 4:06. It introduced direct imports of the Common
JSONL configuration functions into a standalone executable that does not link
that API. Both LNK2019 errors and LNK1120 remain in the unmodified build log.
The driver stopped before launching any native probe. This is a failed build,
not a test result or qualification receipt.

The correction removes those imports and selects the existing governed JSONL
environment in the owned diagnostic launcher before process/plugin startup.
Native queue/event/state observations and all original assertions/deadlines stay
intact. No product export behavior changes in this correction. The corrected
build, stage diagnosis, repair and final qualification remain required.
