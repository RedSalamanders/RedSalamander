# Requalified Debug build evidence

Test-enabled x64 Debug build for HEAD `d713d469213659f20d47f842d6980c1c85b24609`, source snapshot `f43deb78049356ccb5e5fff99ba0f9f90b355b24d62a985b7df124c77ce2aea9`, receipt `2bf8224f970817fc029a478e7bcb4a838e959c32050bb7afbaea2ebdd5468647`.

This is the final Debug build receipt associated with the PrePush run archived in `Validation/space_cache_policy_debug_prepush`. The receipt records exit code 0. `msbuild.log` is the exact corresponding MSBuild transcript; it shows the application output and contains no compiler warning/error diagnostic lines. The PrePush validation itself is FAILED; this archive claims only the successful build.

`manifest.txt` records exact source paths, byte lengths, and SHA-256 values. `.gitattributes` disables text normalization for byte preservation.

