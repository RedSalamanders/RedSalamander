// WARNING: cannot be replaced by "#pragma once" because it is included from .rc files and
// the resource compiler does not reliably support "#pragma once".
#ifndef __REDSALAMANDER_VERSION_H
#define __REDSALAMANDER_VERSION_H

#if defined(APSTUDIO_INVOKED) && ! defined(APSTUDIO_READONLY_SYMBOLS)
#error this file is not editable by Microsoft Visual C++
#endif // defined(APSTUDIO_INVOKED) && !defined(APSTUDIO_READONLY_SYMBOLS)

#ifdef RC_INVOKED
#define VERSINFO_TEXT(s) s
#define VERSINFO_WSTR(s) VERSINFO_STR(s)
#else
#define VERSINFO_TEXT(s) L##s
#endif

#ifdef RS_VERSION_USE_OVERRIDES
#include "RSVersionOverrides.h"
#endif

#define VERSINFO_COPYRIGHT VERSINFO_TEXT("Copyright (c) 2026 RedSalamander Authors")
#define VERSINFO_COMPANY VERSINFO_TEXT("RedSalamander")
#define VERSINFO_PRODUCTNAME VERSINFO_TEXT("RedSalamander")

#define VERSINFO_DESCRIPTION VERSINFO_TEXT("RedSalamander, File Manager")
#define VERSINFO_COMMENT VERSINFO_TEXT("A two-pane file manager with plugin architecture.")

// Human-maintained product line. CI supplies the build number.
//
// Official versions use:
//   <major>.<minor>.<build>
//
// The build number is intentionally not edited here:
// - Local beta builds reuse the saved per-worktree build number under .build\version\ so ordinary
//   incremental builds keep the same version stamp until the caller supplies a new number.
// - CI builds use GITHUB_RUN_NUMBER unless the caller supplies an explicit build number.
// This keeps git clean while making CI the source of truth for release numbering.
#define VERSINFO_MAJOR 7
#define VERSINFO_MINOR 0

// Conversion macros num->str / num->wide-str.
#define VERSINFO_STR_IMPL(s) #s
#define VERSINFO_STR(s) VERSINFO_STR_IMPL(s)
#ifndef RC_INVOKED
#define VERSINFO_WSTR_IMPL(s) L##s
#define VERSINFO_WSTR_EXPAND(s) VERSINFO_WSTR_IMPL(s)
#define VERSINFO_WSTR(s) VERSINFO_WSTR_EXPAND(VERSINFO_STR(s))
#endif
// The build system overrides this value through RSVersionOverrides.h.
// Keep the fallback at 0 so direct single-file compilation still succeeds.
#ifndef VERSINFO_BUILDNUMBER
#define VERSINFO_BUILDNUMBER 0
#endif

// The build system normally injects platform/configuration selectors too. These fallbacks keep
// ad-hoc compilation usable when the project is not launched through build.ps1/MSBuild.
#if ! defined(RSV_PLATFORM_X64) && ! defined(RSV_PLATFORM_ARM64)
#if defined(_M_ARM64)
#define RSV_PLATFORM_ARM64 1
#else
#define RSV_PLATFORM_X64 1
#endif
#endif

#if ! defined(RSV_CFG_DEBUG) && ! defined(RSV_CFG_RELEASE) && ! defined(RSV_CFG_ASAN_DEBUG)
#if defined(_DEBUG)
#define RSV_CFG_DEBUG 1
#else
#define RSV_CFG_RELEASE 1
#endif
#endif

#if defined(RSV_PLATFORM_ARM64)
#define REDSALAMANDER_VER_PLATFORM VERSINFO_TEXT("ARM64")
#else
#define REDSALAMANDER_VER_PLATFORM VERSINFO_TEXT("x64")
#endif

#if defined(RSV_OFFICIAL_RELEASE)
#define VERSINFO_BUILD_FLAVOR_TXT VERSINFO_TEXT("")
#define VERSINFO_BUILD_FLAVOR_TXT_NO_PLATFORM VERSINFO_TEXT("")
#define VERSINFO_BUILD_FLAVOR_SHORT VERSINFO_TEXT("")
#elif defined(RSV_CFG_ASAN_DEBUG)
#define VERSINFO_BUILD_FLAVOR_TXT VERSINFO_TEXT(" beta ASan Debug")
#define VERSINFO_BUILD_FLAVOR_TXT_NO_PLATFORM VERSINFO_TEXT(" beta ASan Debug")
#define VERSINFO_BUILD_FLAVOR_SHORT VERSINFO_TEXT("BAD")
#elif defined(RSV_CFG_DEBUG)
#define VERSINFO_BUILD_FLAVOR_TXT VERSINFO_TEXT(" beta Debug")
#define VERSINFO_BUILD_FLAVOR_TXT_NO_PLATFORM VERSINFO_TEXT(" beta Debug")
#define VERSINFO_BUILD_FLAVOR_SHORT VERSINFO_TEXT("BD")
#else
#define VERSINFO_BUILD_FLAVOR_TXT VERSINFO_TEXT(" beta Release")
#define VERSINFO_BUILD_FLAVOR_TXT_NO_PLATFORM VERSINFO_TEXT(" beta Release")
#define VERSINFO_BUILD_FLAVOR_SHORT VERSINFO_TEXT("BR")
#endif

// Windows VERSIONINFO numeric values require four integer components. The visible FileVersion
// string below stays normalized as <major>.<minor>.<build>.
#define VERSINFO_FILEVERSION_VALUES VERSINFO_MAJOR, VERSINFO_MINOR, VERSINFO_BUILDNUMBER, 0
#define VERSINFO_PRODUCTVERSION_VALUES VERSINFO_MAJOR, VERSINFO_MINOR, VERSINFO_BUILDNUMBER, 0

#define VERSINFO_VERSION_BASE VERSINFO_WSTR(VERSINFO_MAJOR) VERSINFO_TEXT(".") VERSINFO_WSTR(VERSINFO_MINOR)
#define VERSINFO_PACKAGE_VERSION VERSINFO_VERSION_BASE VERSINFO_TEXT(".") VERSINFO_WSTR(VERSINFO_BUILDNUMBER)
#define VERSINFO_REDSALAMANDER_SHORT VERSINFO_WSTR(VERSINFO_MAJOR) VERSINFO_WSTR(VERSINFO_MINOR) VERSINFO_BUILD_FLAVOR_SHORT REDSALAMANDER_VER_PLATFORM

#if defined(RSV_OFFICIAL_RELEASE)
#define VERSINFO_VERSION VERSINFO_PACKAGE_VERSION
#define VERSINFO_VERSION_NO_PLATFORM VERSINFO_PACKAGE_VERSION
#else
#define VERSINFO_VERSION VERSINFO_PACKAGE_VERSION VERSINFO_BUILD_FLAVOR_TXT VERSINFO_TEXT(" (") REDSALAMANDER_VER_PLATFORM VERSINFO_TEXT(")")
#define VERSINFO_VERSION_NO_PLATFORM VERSINFO_PACKAGE_VERSION VERSINFO_BUILD_FLAVOR_TXT_NO_PLATFORM
#endif

#define VERSINFO_REDSALAMANDER VERSINFO_VERSION
#define VERSINFO_PLUGIN_VERSION VERSINFO_VERSION
#define VERSINFO_VERSION_LABEL VERSINFO_TEXT("Version ") VERSINFO_VERSION

// Used to check plugin/host SDK compatibility. The SDK contract now follows the human-maintained
// semantic digits automatically so it cannot drift from the declared product version.
#define LAST_VERSION_OF_SALAMANDER ((VERSINFO_MAJOR * 100) + VERSINFO_MINOR)
#define REQUIRE_LAST_VERSION_OF_REDSALAMANDER                                                                                                                  \
    VERSINFO_TEXT("This plugin requires RedSalamander ") VERSINFO_VERSION_BASE VERSINFO_TEXT(" (") REDSALAMANDER_VER_PLATFORM VERSINFO_TEXT(") or later.")

#endif // __REDSALAMANDER_VERSION_H
