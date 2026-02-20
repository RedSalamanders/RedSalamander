#pragma once
#include "Framework.h"

class Configuration
{
public:
    Configuration();

    bool Save();
    bool Load();

    // Filter settings
    uint32_t filterMask  = 0x1F; // All 5 types enabled by default (bits 0-4)
    int lastFilterPreset = -1;   // -1 = custom, 0 = Errors Only, 1 = Errors+Warnings, 2 = All, 3 = Errors+Debug
};

extern Configuration g_config;
