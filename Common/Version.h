// WARNING: cannot be replaced by "#pragma once" because it is included from .rc file and it seems resource compiler does not support "#pragma once"
#ifndef __REDSALAMANDER_VERSION_H
#define __REDSALAMANDER_VERSION_H

#if defined(APSTUDIO_INVOKED) && ! defined(APSTUDIO_READONLY_SYMBOLS)
#error this file is not editable by Microsoft Visual C++
#endif // defined(APSTUDIO_INVOKED) && !defined(APSTUDIO_READONLY_SYMBOLS)

#define VERSINFO_COPYRIGHT L"Copyright © 2025 Red Salamander Authors"
#define VERSINFO_COMPANY L"Red Salamander"

#define VERSINFO_DESCRIPTION L"Red Salamander, File Manager"
#define VERSINFO_COMMENT L"A two-pane file manager with plugin architecture."

// conversion macros num->str
#define VERSINFO_xstr(s) VERSINFO_str(s)
#define VERSINFO_str(s) #s

#define VERSINFO_MAJOR 7
#define VERSINFO_MINORA 0
#define VERSINFO_MINORB 0

// if minorB is 0, we don't display it in version strings (e.g., 7.50 -> 7.5)
#if (VERSINFO_MINORB == 0)
#define VERSINFO_REDSALAMANDER VERSINFO_xstr(VERSINFO_MAJOR) L"." VERSINFO_xstr(VERSINFO_MINORA) VERSINFO_BETAVERSION_TXT
#define VERSINFO_REDSALAMANDER_SHORT VERSINFO_xstr(VERSINFO_MAJOR) VERSINFO_xstr(VERSINFO_MINORA) VERSINFO_BETAVERSIONSHORT_TXT
#else
#define VERSINFO_REDSALAMANDER VERSINFO_xstr(VERSION_MAJOR) L"." VERSINFO_xstr(VERSINFO_MINORA) VERSINFO_xstr(VERSINFO_MINORB) VERSINFO_BETAVERSION_TXT
#define VERSINFO_REDSALAMANDER_SHORT VERSINFO_xstr(VERSION_MAJOR) VERSINFO_xstr(VERSINFO_MINORA) VERSINFO_xstr(VERSINFO_MINORB) VERSINFO_BETAVERSIONSHORT_TXT
#endif

#ifdef VERSINFO_MAJOR      // defined only when used from a plugin
#if (VERSINFO_MINORB == 0) // if the hundredths are zero, we don't show them (e.g., 2.50 -> 2.5)
#define VERSINFO_VERSION VERSINFO_xstr(VERSINFO_MAJOR) L"." VERSINFO_xstr(VERSINFO_MINORA) VERSINFO_BETAVERSION_TXT
#define VERSINFO_VERSION_NO_PLATFORM VERSINFO_xstr(VERSINFO_MAJOR) L"." VERSINFO_xstr(VERSINFO_MINORA) VERSINFO_BETAVERSION_TXT_NO_PLATFORM
#else
#define VERSINFO_VERSION VERSINFO_xstr(VERSINFO_MAJOR) "." VERSINFO_xstr(VERSINFO_MINORA) VERSINFO_xstr(VERSINFO_MINORB) VERSINFO_BETAVERSION_TXT
#define VERSINFO_VERSION_NO_PLATFORM                                                                                                                           \
    VERSINFO_xstr(VERSINFO_MAJOR) "." VERSINFO_xstr(VERSINFO_MINORA) VERSINFO_xstr(VERSINFO_MINORB) VERSINFO_BETAVERSION_TXT_NO_PLATFORM
#endif
#endif

#define REDSALAMANDER_VER_PLATFORM L"x64"

// VERSINFO_BUILDNUMBER:
//
// Used to easily distinguish versions of all modules across releases
// (this is the last component of the version number for all plugins and RedSalamander itself).
// Increment with every version (IB, DB, PB, beta, release, or even a test version sent to a single user).
// An overview of version types is in doc\versions.txt.
// Always add a comment explaining which RedSalamander version the new build number corresponds to.
//
// Overview of used VERSINFO_BUILDNUMBER values:
//...
// ! IMPORTANT: new build numbers must be added to the "default" branch first,
//              and only then to side branches (the complete list exists only in the "default" branch)
#define VERSINFO_BUILDNUMBER 183

// VERSINFO_BETAVERSION_TXT:
//
// Changes with every build; for release versions, VERSINFO_BETAVERSION_TXT = "".
// If releasing a special fix beta version like "2.5 beta 9a", increment
// VERSINFO_BUILDNUMBER by one and set VERSINFO_BETAVERSION_TXT == " beta 9a".
//
// VERSINFO_BETAVERSIONSHORT_TXT is used for naming bug reports; it should be as short as possible

// examples ("x64" for 64-bit builds; interchangeable in examples below):
// " beta 2 (x64)", " beta 2 (SDK xArm)",
// " RC1 (x64)", " beta 2 (IB21 xArm)", " beta 2 (DB21 x64)", " beta 2 (PB21 xArm)"
#define VERSINFO_BETAVERSION_TXT L" (" REDSALAMANDER_VER_PLATFORM L")"
#define VERSINFO_BETAVERSION_TXT_NO_PLATFORM                                                                                                                   \
    L"" // copy the line above + remove REDSALAMANDER_VER_PLATFORM + if parentheses are empty, remove them + remove extra spaces

// examples (see above for x64): "x64" (for release), "B2x64", "B2SDKx64",
// "RC1x64", "B2IB21x64", "B2DB21x64", "B2PB21x64"
#define VERSINFO_BETAVERSIONSHORT_TXT REDSALAMANDER_VER_PLATFORM

// LAST_VERSION_OF_REDSALAMANDER:
//
// Used to check the compatibility of RedSalamander plugins during their entry point
// (see PluginEntryAbstract::GetVersion() in plugin_base.h).
// Mainly serves simplicity: internal plugins can call any method from RedSalamander's interface,
// because after checking for this version, they are guaranteed it is supported by RedSalamander.
// (Only a newer RedSalamander version might load them, which must also include these methods.)
//
// Also used in reverse: to ensure RedSalamander will call all methods of a plugin (including the newest),
// the plugin returns this version via the PluginGetRequiredVersion export.
//
// If a plugin returns a lower version from PluginGetRequiredVersion (for backward compatibility),
// it should add the PluginGetSDKVersion export and return LAST_VERSION_OF_SALAMANDER
// to indicate which SDK version was used for compilation—so that RedSalamander (e.g., newer version)
// can use methods from the plugin not present in older versions.
//
// When changing the interface, follow the procedure described in doc\how_to_change.txt.
//
#define LAST_VERSION_OF_SALAMANDER 703
#define REQUIRE_LAST_VERSION_OF_REDSALAMANDER L"This plugin requires Red Salamander 7.0 (" REDSALAMANDER_VER_PLATFORM L") or later."

#endif // __SPL_VERS_H
