# MSVC 19.52.37025 ARM64 ICE in lookup3.inl while compiling hash_table.c.
# Preserve /O2 and explicit-inline expansion; restrict only opportunistic inlining
# in this one third-party translation unit. Other files/architectures/tools are unchanged.
include_guard(GLOBAL)
if(MSVC AND CMAKE_C_COMPILER_VERSION VERSION_EQUAL "19.52.37025.0" AND CMAKE_C_COMPILER_ARCHITECTURE_ID STREQUAL "ARM64")
    set_property(SOURCE "${CMAKE_CURRENT_SOURCE_DIR}/source/hash_table.c" APPEND PROPERTY COMPILE_OPTIONS "/Ob1")
endif()
