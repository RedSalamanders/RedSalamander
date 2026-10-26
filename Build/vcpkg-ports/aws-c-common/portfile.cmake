vcpkg_from_github(
    OUT_SOURCE_PATH SOURCE_PATH
    REPO awslabs/aws-c-common
    REF "v${VERSION}"
    SHA512 f9cd67637383d86353cef7d230051aad67574f9b5c72d6260713f9d8b1139d730ce56d3cc8b646bc0344d99c8cae79d3cfca86c1994207c1b74b9bb814645f4b
    HEAD_REF master
    PATCHES
        disable-internal-crt-option.patch # Disable internal crt option because vcpkg contains crt processing flow
)

# Local compatibility overlay of registry tree a4b4f45d42ae3a7ce2c5040af18cc9bb1b0113a6.
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
        "${SOURCE_PATH}/THIRD-PARTY-LICENSES.txt"
)
