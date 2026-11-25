# I66 Create Folder admission after accepted close

The test-enabled x64 Debug application built with zero warnings and zero errors. Run
`20261124T190046Z-93952-092a52e3388f44c386698349324f41e8` used receipt
`237be47fb41f20903bf47e6d404cce30fa6d639de93287af850e3c85c3aabb10`.
It passed setup, `FileOps_DeferredExitRefusesNewWork` and cleanup without desktop activation.
The focused admission case took 2,047 ms. The scope result was 3 passed, 0 failed and 0 skipped.
Repository coverage remained unevaluated.

After an accepted exit was waiting for existing work to stop, Create Folder showed the localized
refusal in its pane. It opened zero name or consent prompts and submitted zero tasks. The fixture
files and queue stayed intact. The same case verified the other ingress refusals, including a
canceled drag-and-drop with no reported transfer effect.

The preceding run `20261124T184628Z-96656-cbfc3e3711c5438c8a5394277f1df50f` verified the
Create Folder refusal and failed at its stale drag-and-drop status assertion. The current test checks
the actual cancellation status and the absence of a transfer effect.

The complete current trace has 18,758 bytes and SHA-256
`70e81845ea1dd8ace245165611c537be0429b4449425bd19eaae6295a0263ff4`.
The unified sandbox run keeps the complete trace and results.
The async Create Folder path and its latency witnesses remain open.
