# Throughput disclosure setting candidate

This candidate implements `fileOperations.popupThroughputExpanded` as a typed boolean setting with
an expanded-by-default value. Explicit `false` parses and survives canonical save preparation; absent,
explicit `true`, and invalid-type values resolve to `true` and are pruned when otherwise default.

The x64 Debug build completed with zero warnings and zero errors. Its receipt is
`3de6b88af79bfc0278a990ff0de4666bd71f70180ae6bc815d132ba83d36e016`, bound to source snapshot
`4fc5a2ec0cb3759bb86be234d86d79fe667ddeefc216e919b250efddc1f4ce67`.
The governed focused Commands run `20260919T133051Z-76472-2b4998908dba40dcbf2937cadee166cd`
passed 1/1 with zero skips and a zero-issue sandbox audit. The retained result, trace, runner log,
performance stream, and aggregate result preserve that verdict.

Documentation drift contracts passed 18/18, and the specification inventory reported zero blocking
findings. Two direct launches of `SettingsSchemaTests.exe`, including one with isolated local app data,
failed unrelated SettingsStore CAS/read-sharing environment guards and terminated with
`-1073741819`; those launches are diagnostic and are not claimed as a standalone pass. The required
final governed Full run remains the closeout authority.

This evidence qualifies the storage contract only. The disclosure control, live-card synchronization,
qualified DxUi pin adoption, native mixed-DPI acceptance, and final Full run remain open.
