# Preserved final Debug preflight failure

RedSalamander `ad7f7735` passes the local x64 Debug root build with zero warnings/errors.
Fresh Full then exits 3 before test execution, run
`20260913T004508Z-2916-cfd52ac0b56c49928b7913557d0b0489`, because the diagnostic
`before-DxUi-resolved-ARM64.props` snapshot has an MSBuild suffix but is not imported by the
active project graph. The validator and active graph are unchanged; the snapshot now has a
`.props.txt` suffix, retaining the original bytes and manifest source path. This is a corrected
evidence filename, not a test waiver. Build receipt:
`19fee914abe1320b5a562dce199a00e3d3e5f62e278b9c4d7de9f8bdf2cd3d38`.
The final three-profile sequence must be restarted after this documentation/evidence change.
