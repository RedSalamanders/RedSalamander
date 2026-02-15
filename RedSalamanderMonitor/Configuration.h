#pragma once
#include "Framework.h"

#define DESCRIPTION_SIZE 5000
#define EMAIL_SIZE 100

class Configuration
{
public:
    Configuration();

    BOOL Save();
    BOOL Load();

    std::wstring description;
    std::wstring email;
    BOOL restart = TRUE; // do not save

    // Filter settings
    uint32_t filterMask  = 0x1F; // All 5 types enabled by default (bits 0-4)
    int lastFilterPreset = -1;   // -1 = custom, 0 = Errors Only, 1 = Errors+Warnings, 2 = All, 3 = Errors+Debug
};

extern Configuration g_config;
