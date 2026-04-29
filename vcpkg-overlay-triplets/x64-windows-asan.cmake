set(VCPKG_TARGET_ARCHITECTURE x64)
set(VCPKG_CRT_LINKAGE dynamic)
set(VCPKG_LIBRARY_LINKAGE dynamic)

# Build native dependencies with /fsanitize=address so the MSVC STL emits
# matching annotate_optional / annotate_string / annotate_vector link tokens.
# Without this, consuming first-party projects in the "ASan Debug" solution
# configuration cannot link against vcpkg-provided static libs (LNK2038).
set(VCPKG_C_FLAGS "/fsanitize=address")
set(VCPKG_CXX_FLAGS "/fsanitize=address")

# ASan requires /INCREMENTAL:NO. We prepend it here; OpenSSL's portfile.cmake
# will strip any conflicting /INCREMENTAL from vcpkg-detected MSVC defaults.
set(VCPKG_LINKER_FLAGS "/INCREMENTAL:NO")
