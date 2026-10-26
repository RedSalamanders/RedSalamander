# Debug full-build evidence

This archive preserves the test-enabled x64 Debug build completed for source snapshot `f43deb78049356ccb5e5fff99ba0f9f90b355b24d62a985b7df124c77ce2aea9` at HEAD `d713d469213659f20d47f842d6980c1c85b24609`.

The immutable receipt records exit code 0. The exact MSBuild transcript is retained; it shows the application output and contains no compiler warning/error diagnostic lines. These files establish the build only; the associated PrePush validation was still running when this snapshot was archived, so this archive does not claim PrePush completion or pass.

`manifest.txt` records SHA-256 and byte length for each preserved source copy. `.gitattributes` disables text conversion for all archive files so Git preserves their bytes.

