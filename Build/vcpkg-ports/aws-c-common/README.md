# AWS common compiler compatibility overlay

Upstream registry tree a4b4f45d42ae3a7ce2c5040af18cc9bb1b0113a6, aws-c-common 1.0.0. Keep its source SHA512 and CRT patch unchanged. Only ARM64 dependency installs select this overlay. Its CMake helper limits automatic inlining to /Ob1 in source/hash_table.c under the exact MSVC C compiler 19.52.37025.0; /O2 remains enabled and other files/architectures/compilers are unaffected.

The exact 14.52.37025 ARM64 command deterministically reports C1001 in lookup3.inl:640. Removing /MP or changing /Z7 to /Zi still crashes; /O1 also crashes. /Ob1 and /Ob0 compile successfully. Original/reproduction logs are retained under .build/msvc1452/failed-aws-arm64 and compiler-repro-attempt2/3. The canonical dependency and product builds must still pass. ARM64 runtime/S3 hash performance requires native execution; this cross-build workaround does not qualify it.

The packaging files derive from Microsoft vcpkg under its MIT license (LICENSE.vcpkg-MIT.txt); upstream AWS source is downloaded unchanged under its own package licenses.
