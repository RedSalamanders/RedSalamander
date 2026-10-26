# Terminal ARM64 product dependency preflight

Command:

```powershell
.\build.ps1 -ProjectName Terminal -Configuration Debug -Platform ARM64
```

Result: build wrapper exit 1, log
`.build/logs/msbuild-20260802_222051_655.log`, 0 warnings and 2 wrapper errors.
The failure occurred while vcpkg configured `7zip:arm64-windows`, before any
RedSalamander or Terminal translation unit compiled.

The CMake probe identified target system ARM64 but selected
`MSVC/14.52.36314/bin/Hostx64/x64/cl.exe`, generated `/machine:x64`, and failed
its trivial link with `LNK1104` for `MSVCRTD.lib` and `LIBCMT.lib`. The pinned
Ghostty ARM64 runtime remains independently qualified as ARM64 with only
`kernel32.dll` and `ntdll.dll` imports and SHA-256
`b32b4854d6d57494576c2579c13582a84ab21ef1efd82c7ffe7c287fca6e2e75`.

The failed cross-platform vcpkg transaction removed the shared x64 installed
set before stopping. Closeout therefore reinstalled all 32 pinned x64 manifest
packages, rebuilt Terminal successfully in
`.build/logs/msbuild-20260802_224323_251.log` with 0 warnings/0 errors, and
reran the real terminal lifecycle 10/10 under
`20260802T204347Z-2000-ced799a18ae64b8ea52d28e82e290865`.

Disposition: release toolchain/dependency gate remains open. Repair the Visual
Studio v145/vcpkg ARM64 environment, rerun the same command, then run the native
ARM64 product/test/package lanes. Do not weaken the product platform requirement
or count the engine-only ARM64 artifact as a product build.
