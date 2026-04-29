#pragma once
#include "Framework.h"

class Configuration
{
public:
    Configuration();

    bool Save();
    bool Load();

    // Filter settings
    uint32_t filterMask  = 0x3F; // All 6 visible types enabled by default (bits 0-5)
    int lastFilterPreset = -1;   // -1 = custom, 0 = Errors Only, 1 = Errors+Warnings, 2 = All, 3 = Errors+Perf+Debug
};

extern Configuration g_config;
