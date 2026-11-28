# Extend the Windows toolchain of the selected vcpkg checkout. The fallback
# is relative to the generated overlay in .build/vcpkg-triplets/<platform>.
if(Z_VCPKG_ROOT_DIR)
    list(APPEND CMAKE_TRY_COMPILE_PLATFORM_VARIABLES Z_VCPKG_ROOT_DIR)
    include("${Z_VCPKG_ROOT_DIR}/scripts/toolchains/windows.cmake")
else()
    include("${CMAKE_CURRENT_LIST_DIR}/../../vcpkg-tool/scripts/toolchains/windows.cmake")
endif()

# MSVC rejects /RTC1 when compiling with /fsanitize=address.
foreach(lang C CXX)
    string(REPLACE "/RTC1" "" debug_flags "${CMAKE_${lang}_FLAGS_DEBUG}")
    set(CMAKE_${lang}_FLAGS_DEBUG "${debug_flags}" CACHE STRING "" FORCE)
endforeach()
unset(debug_flags)
