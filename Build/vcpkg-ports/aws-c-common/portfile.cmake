vcpkg_from_github(
    OUT_SOURCE_PATH SOURCE_PATH
    REPO awslabs/aws-c-common
    REF "v${VERSION}"
    SHA512 05cd13aa8c55762345fd7754c1c393b6db48918cb4b4335ade9a75dfbcf20802fa00f8d2e25fa2826b054ffdfe98c60bc50f6c7bb686c1d7df790741bbf8244a
    HEAD_REF master
    PATCHES
        disable-internal-crt-option.patch # Disable internal crt option because vcpkg contains crt processing flow
)

# Local compatibility overlay of registry tree 97972a5d5632f2cf4a5d649e8b261f89059643bc.
# Keep the upstream source/version/hash and CRT patch. The helper checks the actual
# compiler/target before applying the narrowly reproduced ARM64 backend workaround.
file(APPEND "${SOURCE_PATH}/CMakeLists.txt" "\ninclude(\"${CMAKE_CURRENT_LIST_DIR}/msvc1452-arm64.cmake\")\n")
vcpkg_cmake_configure(
    SOURCE_PATH "${SOURCE_PATH}"
    OPTIONS
        -DBUILD_TESTING=FALSE
)

vcpkg_cmake_install()

string(REPLACE "dynamic" "shared" subdir "${VCPKG_LIBRARY_LINKAGE}")
vcpkg_cmake_config_fixup(CONFIG_PATH "lib/cmake/${PORT}/${subdir}" DO_NOT_DELETE_PARENT_CONFIG_PATH)
vcpkg_cmake_config_fixup(CONFIG_PATH "lib/cmake/${PORT}")
vcpkg_replace_string("${CURRENT_PACKAGES_DIR}/share/${PORT}/${PORT}-config.cmake" [[/${type}/]] "/")

file(REMOVE_RECURSE
    "${CURRENT_PACKAGES_DIR}/debug/include"
    "${CURRENT_PACKAGES_DIR}/debug/lib/${PORT}"
	"${CURRENT_PACKAGES_DIR}/debug/share"
    "${CURRENT_PACKAGES_DIR}/lib/${PORT}"
)

vcpkg_copy_pdbs()

vcpkg_install_copyright(
    FILE_LIST
        "${SOURCE_PATH}/LICENSE"
        "${SOURCE_PATH}/NOTICE"
        "${SOURCE_PATH}/THIRD-PARTY-LICENSES.txt"
)
