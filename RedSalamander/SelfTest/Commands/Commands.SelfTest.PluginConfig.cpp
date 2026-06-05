// Commands.SelfTest.PluginConfig.cpp
// Included from Commands.SelfTest.cpp — NOT compiled standalone.
// PluginConfig test family: 14 test functions.

[[nodiscard]] bool TestFileSystemPluginConfigurationRoundTrip(CaseState& state) noexcept
{
    CreatedFileSystemInstance created{};
    const HRESULT createHr = TryCreateFileSystemInstance(kBuiltinLocalFileSystemId, created);
    state.Require(SUCCEEDED(createHr) && created.fileSystem,
                  std::format(L"Failed to create isolated local file system instance. hr=0x{:08X}", static_cast<unsigned long>(createHr)));
    if (FAILED(createHr) || ! created.fileSystem)
    {
        return false;
    }

    wil::com_ptr<IInformations> info;
    state.Require(CreateInformations(created.fileSystem, info), L"Isolated local file system instance missing IInformations.");
    if (! info)
    {
        return false;
    }

    const char* schema     = nullptr;
    const HRESULT schemaHr = info->GetConfigurationSchema(&schema);
    state.Require(SUCCEEDED(schemaHr) && schema != nullptr,
                  std::format(L"GetConfigurationSchema failed for the local file system plugin. hr=0x{:08X}", static_cast<unsigned long>(schemaHr)));
    if (schema != nullptr)
    {
        const std::string_view schemaView(schema);
        state.Require(schemaView.find("\"recycleBinBatchSize\"") != std::string_view::npos, L"Configuration schema missing recycleBinBatchSize.");
        state.Require(schemaView.find("\"searchMaxDirectoryWalkers\"") != std::string_view::npos, L"Configuration schema missing searchMaxDirectoryWalkers.");
        state.Require(schemaView.find("\"copyMoveMaxConcurrency\"") != std::string_view::npos, L"Configuration schema missing copyMoveMaxConcurrency.");
        state.Require(schemaView.find("\"concurrencyMode\"") != std::string_view::npos, L"Configuration schema missing concurrencyMode.");
        state.Require(schemaView.find("\"default\": 500") != std::string_view::npos, L"Configuration schema missing recycleBinBatchSize default.");
        state.Require(schemaView.find("\"default\": 4") != std::string_view::npos, L"Configuration schema missing searchMaxDirectoryWalkers default.");
        state.Require(schemaView.find("\"default\": \"auto\"") != std::string_view::npos, L"Configuration schema missing concurrencyMode default.");
        state.Require(schemaView.find("\"max\": 1000") != std::string_view::npos, L"Configuration schema missing recycleBinBatchSize max.");
        state.Require(schemaView.find("\"max\": 8") != std::string_view::npos, L"Configuration schema missing searchMaxDirectoryWalkers max.");
        state.Require(schemaView.find("\"max\": 16") != std::string_view::npos, L"Configuration schema missing copyMoveMaxConcurrency max.");
    }

    const std::string resetConfiguration =
        R"json({"concurrencyMode":"auto","copyMoveMaxConcurrency":4,"searchBackendPreference":"auto","recycleBinBatchSize":500,"searchMaxDirectoryWalkers":4})json";
    const HRESULT setDefaultHr = info->SetConfiguration(resetConfiguration.c_str());
    state.Require(
        SUCCEEDED(setDefaultHr),
        std::format(L"SetConfiguration failed when seeding default local file system plugin settings. hr=0x{:08X}", static_cast<unsigned long>(setDefaultHr)));
    if (FAILED(setDefaultHr))
    {
        return false;
    }

    BOOL somethingToSave        = TRUE;
    const HRESULT defaultSaveHr = info->SomethingToSave(&somethingToSave);
    state.Require(
        SUCCEEDED(defaultSaveHr),
        std::format(L"SomethingToSave failed for default local file system plugin configuration. hr=0x{:08X}", static_cast<unsigned long>(defaultSaveHr)));
    if (SUCCEEDED(defaultSaveHr))
    {
        state.Require(somethingToSave == FALSE, L"Default local file system plugin configuration should not be dirty.");
    }

    const std::string customConfiguration =
        R"json({"concurrencyMode":"manual","copyMoveMaxConcurrency":12,"searchBackendPreference":"scan","recycleBinBatchSize":256,"searchMaxDirectoryWalkers":6})json";
    const HRESULT setCustomHr = info->SetConfiguration(customConfiguration.c_str());
    state.Require(SUCCEEDED(setCustomHr),
                  std::format(L"SetConfiguration failed for custom local file system plugin settings. hr=0x{:08X}", static_cast<unsigned long>(setCustomHr)));
    if (FAILED(setCustomHr))
    {
        return false;
    }

    somethingToSave            = FALSE;
    const HRESULT customSaveHr = info->SomethingToSave(&somethingToSave);
    state.Require(SUCCEEDED(customSaveHr),
                  std::format(L"SomethingToSave failed after custom local file system plugin settings. hr=0x{:08X}", static_cast<unsigned long>(customSaveHr)));
    if (SUCCEEDED(customSaveHr))
    {
        state.Require(somethingToSave == TRUE, L"Custom local file system plugin settings should be dirty.");
    }

    const char* configuration = nullptr;
    const HRESULT getCustomHr = info->GetConfiguration(&configuration);
    state.Require(SUCCEEDED(getCustomHr) && configuration != nullptr,
                  std::format(L"GetConfiguration failed after custom local file system plugin settings. hr=0x{:08X}", static_cast<unsigned long>(getCustomHr)));
    if (configuration != nullptr)
    {
        const std::string_view configurationView(configuration);
        state.Require(configurationView.find("\"copyMoveMaxConcurrency\":12") != std::string_view::npos,
                      L"Configuration JSON missing copyMoveMaxConcurrency=12.");
        state.Require(configurationView.find("\"concurrencyMode\":\"manual\"") != std::string_view::npos,
                      L"Configuration JSON missing concurrencyMode=manual.");
        state.Require(configurationView.find("\"searchBackendPreference\":\"scan\"") != std::string_view::npos,
                      L"Configuration JSON missing searchBackendPreference=scan.");
        state.Require(configurationView.find("\"recycleBinBatchSize\":256") != std::string_view::npos, L"Configuration JSON missing recycleBinBatchSize=256.");
        state.Require(configurationView.find("\"searchMaxDirectoryWalkers\":6") != std::string_view::npos,
                      L"Configuration JSON missing searchMaxDirectoryWalkers=6.");
    }

    const HRESULT resetHr = info->SetConfiguration(resetConfiguration.c_str());
    state.Require(SUCCEEDED(resetHr),
                  std::format(L"SetConfiguration failed when resetting local file system plugin settings. hr=0x{:08X}", static_cast<unsigned long>(resetHr)));
    if (FAILED(resetHr))
    {
        return false;
    }

    somethingToSave           = TRUE;
    const HRESULT resetSaveHr = info->SomethingToSave(&somethingToSave);
    state.Require(
        SUCCEEDED(resetSaveHr),
        std::format(L"SomethingToSave failed after resetting local file system plugin settings. hr=0x{:08X}", static_cast<unsigned long>(resetSaveHr)));
    if (SUCCEEDED(resetSaveHr))
    {
        state.Require(somethingToSave == FALSE, L"Reset local file system plugin settings should return to a clean state.");
    }

    configuration            = nullptr;
    const HRESULT getResetHr = info->GetConfiguration(&configuration);
    state.Require(
        SUCCEEDED(getResetHr) && configuration != nullptr,
        std::format(L"GetConfiguration failed after resetting local file system plugin settings. hr=0x{:08X}", static_cast<unsigned long>(getResetHr)));
    if (configuration != nullptr)
    {
        const std::string_view configurationView(configuration);
        state.Require(configurationView.find("\"concurrencyMode\":\"auto\"") != std::string_view::npos,
                      L"Configuration JSON missing concurrencyMode=auto after reset.");
        state.Require(configurationView.find("\"searchBackendPreference\":\"auto\"") != std::string_view::npos,
                      L"Configuration JSON missing searchBackendPreference=auto after reset.");
        state.Require(configurationView.find("\"recycleBinBatchSize\":500") != std::string_view::npos,
                      L"Configuration JSON missing recycleBinBatchSize=500 after reset.");
        state.Require(configurationView.find("\"searchMaxDirectoryWalkers\":4") != std::string_view::npos,
                      L"Configuration JSON missing searchMaxDirectoryWalkers=4 after reset.");
    }

    return state.failure.empty();
}

[[nodiscard]] bool TestViewerTextPluginConfigurationRoundTrip(CaseState& state) noexcept
{
    Common::Settings::Settings isolatedSettings{};
    wil::com_ptr<IViewer> viewer;
    const HRESULT createHr = ViewerPluginManager::GetInstance().CreateViewerInstance(L"builtin/viewer-text", isolatedSettings, viewer);
    state.Require(SUCCEEDED(createHr) && viewer,
                  std::format(L"Failed to create isolated ViewerText instance. hr=0x{:08X}", static_cast<unsigned long>(createHr)));
    if (FAILED(createHr) || ! viewer)
    {
        return false;
    }

    wil::com_ptr<IInformations> info;
    const HRESULT infoHr = viewer->QueryInterface(__uuidof(IInformations), info.put_void());
    state.Require(SUCCEEDED(infoHr) && info, std::format(L"ViewerText instance missing IInformations. hr=0x{:08X}", static_cast<unsigned long>(infoHr)));
    if (FAILED(infoHr) || ! info)
    {
        return false;
    }

    const char* schema     = nullptr;
    const HRESULT schemaHr = info->GetConfigurationSchema(&schema);
    state.Require(SUCCEEDED(schemaHr) && schema != nullptr,
                  std::format(L"GetConfigurationSchema failed for ViewerText. hr=0x{:08X}", static_cast<unsigned long>(schemaHr)));
    if (schema != nullptr)
    {
        const std::string_view schemaView(schema);
        state.Require(schemaView.find("\"hexByteColorMode\"") != std::string_view::npos, L"ViewerText configuration schema missing hexByteColorMode.");
        state.Require(schemaView.find("\"default\": \"leadingNibble\"") != std::string_view::npos,
                      L"ViewerText configuration schema missing hexByteColorMode default=leadingNibble.");
        state.Require(schemaView.find("\"leadingNibble\"") != std::string_view::npos, L"ViewerText configuration schema missing leadingNibble option.");
        state.Require(schemaView.find("\"diffDefaultLayout\"") != std::string_view::npos, L"ViewerText configuration schema missing diffDefaultLayout.");
        state.Require(schemaView.find("\"diffContextMode\"") != std::string_view::npos, L"ViewerText configuration schema missing diffContextMode.");
        state.Require(schemaView.find("\"diffAutoOpenMode\"") != std::string_view::npos, L"ViewerText configuration schema missing diffAutoOpenMode.");
        state.Require(schemaView.find("\"diffVisualStyle\"") == std::string_view::npos,
                      L"ViewerText configuration schema should not introduce a persisted diffVisualStyle knob.");
        state.Require(schemaView.find("\"diffContextPresentation\"") == std::string_view::npos,
                      L"ViewerText configuration schema should not introduce a persisted diffContextPresentation knob.");
    }

    BOOL somethingToSave        = TRUE;
    const HRESULT defaultSaveHr = info->SomethingToSave(&somethingToSave);
    state.Require(SUCCEEDED(defaultSaveHr),
                  std::format(L"SomethingToSave failed for default ViewerText configuration. hr=0x{:08X}", static_cast<unsigned long>(defaultSaveHr)));
    if (SUCCEEDED(defaultSaveHr))
    {
        state.Require(somethingToSave == FALSE, L"Default ViewerText configuration should not be dirty.");
    }

    const char* configuration  = nullptr;
    const HRESULT getDefaultHr = info->GetConfiguration(&configuration);
    state.Require(SUCCEEDED(getDefaultHr) && configuration != nullptr,
                  std::format(L"GetConfiguration failed for the default ViewerText settings. hr=0x{:08X}", static_cast<unsigned long>(getDefaultHr)));
    if (configuration != nullptr)
    {
        const std::string_view configurationView(configuration);
        state.Require(configurationView.find("\"hexByteColorMode\":\"leadingNibble\"") != std::string_view::npos,
                      L"Default ViewerText configuration JSON missing hexByteColorMode=leadingNibble.");
        state.Require(configurationView.find("\"diffDefaultLayout\":\"sideBySide\"") != std::string_view::npos,
                      L"Default ViewerText configuration JSON missing diffDefaultLayout=sideBySide.");
        state.Require(configurationView.find("\"diffContextMode\":\"hunksOnly\"") != std::string_view::npos,
                      L"Default ViewerText configuration JSON missing diffContextMode=hunksOnly.");
        state.Require(configurationView.find("\"diffAutoOpenMode\":\"parsed\"") != std::string_view::npos,
                      L"Default ViewerText configuration JSON missing diffAutoOpenMode=parsed.");
        state.Require(configurationView.find("\"diffVisualStyle\"") == std::string_view::npos,
                      L"Default ViewerText configuration JSON should not persist diffVisualStyle.");
        state.Require(configurationView.find("\"diffContextPresentation\"") == std::string_view::npos,
                      L"Default ViewerText configuration JSON should not persist diffContextPresentation.");
    }

    const std::string missingKeyConfiguration = R"json({"textBufferMiB":16,"hexBufferMiB":8,"showLineNumbers":"0","wrapText":"1"})json";
    const HRESULT setMissingKeyHr             = info->SetConfiguration(missingKeyConfiguration.c_str());
    state.Require(
        SUCCEEDED(setMissingKeyHr),
        std::format(L"SetConfiguration failed for ViewerText settings without hexByteColorMode. hr=0x{:08X}", static_cast<unsigned long>(setMissingKeyHr)));
    if (FAILED(setMissingKeyHr))
    {
        return false;
    }

    somethingToSave                = TRUE;
    const HRESULT missingKeySaveHr = info->SomethingToSave(&somethingToSave);
    state.Require(
        SUCCEEDED(missingKeySaveHr),
        std::format(L"SomethingToSave failed after ViewerText settings without hexByteColorMode. hr=0x{:08X}", static_cast<unsigned long>(missingKeySaveHr)));
    if (SUCCEEDED(missingKeySaveHr))
    {
        state.Require(somethingToSave == FALSE, L"ViewerText settings without hexByteColorMode should fall back to the clean default state.");
    }

    configuration                 = nullptr;
    const HRESULT getMissingKeyHr = info->GetConfiguration(&configuration);
    state.Require(
        SUCCEEDED(getMissingKeyHr) && configuration != nullptr,
        std::format(L"GetConfiguration failed after ViewerText settings without hexByteColorMode. hr=0x{:08X}", static_cast<unsigned long>(getMissingKeyHr)));
    if (configuration != nullptr)
    {
        const std::string_view configurationView(configuration);
        state.Require(configurationView.find("\"hexByteColorMode\":\"leadingNibble\"") != std::string_view::npos,
                      L"ViewerText configuration JSON should restore hexByteColorMode=leadingNibble when the key is absent.");
        state.Require(configurationView.find("\"diffDefaultLayout\":\"sideBySide\"") != std::string_view::npos,
                      L"ViewerText configuration JSON should restore diffDefaultLayout=sideBySide when the key is absent.");
        state.Require(configurationView.find("\"diffContextMode\":\"hunksOnly\"") != std::string_view::npos,
                      L"ViewerText configuration JSON should restore diffContextMode=hunksOnly when the key is absent.");
        state.Require(configurationView.find("\"diffAutoOpenMode\":\"parsed\"") != std::string_view::npos,
                      L"ViewerText configuration JSON should restore diffAutoOpenMode=parsed when the key is absent.");
    }

    const std::string transientConfiguration =
        R"json({"textBufferMiB":16,"hexBufferMiB":8,"showLineNumbers":"0","wrapText":"1","hexByteColorMode":"leadingNibble","diffDefaultLayout":"sideBySide","diffContextMode":"hunksOnly","diffAutoOpenMode":"parsed","diffVisualStyle":"semanticRows","diffContextPresentation":"banner","activeDiffHunkIndex":2,"activeDiffSectionIndex":1})json";
    const HRESULT setTransientHr = info->SetConfiguration(transientConfiguration.c_str());
    state.Require(
        SUCCEEDED(setTransientHr),
        std::format(L"SetConfiguration failed for ViewerText settings with transient diff UI state. hr=0x{:08X}", static_cast<unsigned long>(setTransientHr)));
    if (FAILED(setTransientHr))
    {
        return false;
    }

    somethingToSave               = TRUE;
    const HRESULT transientSaveHr = info->SomethingToSave(&somethingToSave);
    state.Require(SUCCEEDED(transientSaveHr),
                  std::format(L"SomethingToSave failed after ViewerText settings with transient diff UI state. hr=0x{:08X}",
                              static_cast<unsigned long>(transientSaveHr)));
    if (SUCCEEDED(transientSaveHr))
    {
        state.Require(somethingToSave == FALSE, L"ViewerText should ignore transient diff UI state instead of treating it as persisted configuration.");
    }

    configuration                = nullptr;
    const HRESULT getTransientHr = info->GetConfiguration(&configuration);
    state.Require(SUCCEEDED(getTransientHr) && configuration != nullptr,
                  std::format(L"GetConfiguration failed after ViewerText settings with transient diff UI state. hr=0x{:08X}",
                              static_cast<unsigned long>(getTransientHr)));
    if (configuration != nullptr)
    {
        const std::string_view configurationView(configuration);
        state.Require(configurationView.find("\"diffVisualStyle\"") == std::string_view::npos,
                      L"ViewerText configuration JSON should not echo transient diffVisualStyle input.");
        state.Require(configurationView.find("\"diffContextPresentation\"") == std::string_view::npos,
                      L"ViewerText configuration JSON should not echo transient diffContextPresentation input.");
        state.Require(configurationView.find("\"activeDiffHunkIndex\"") == std::string_view::npos,
                      L"ViewerText configuration JSON should not persist the active diff hunk index.");
        state.Require(configurationView.find("\"activeDiffSectionIndex\"") == std::string_view::npos,
                      L"ViewerText configuration JSON should not persist the active diff section index.");
    }

    const std::string customConfiguration =
        R"json({"textBufferMiB":16,"hexBufferMiB":8,"showLineNumbers":"0","wrapText":"1","hexByteColorMode":"off","diffDefaultLayout":"inline","diffContextMode":"fullFileWhenAvailable","diffAutoOpenMode":"rawText"})json";
    const HRESULT setCustomHr = info->SetConfiguration(customConfiguration.c_str());
    state.Require(SUCCEEDED(setCustomHr),
                  std::format(L"SetConfiguration failed for custom ViewerText settings. hr=0x{:08X}", static_cast<unsigned long>(setCustomHr)));
    if (FAILED(setCustomHr))
    {
        return false;
    }

    somethingToSave            = FALSE;
    const HRESULT customSaveHr = info->SomethingToSave(&somethingToSave);
    state.Require(SUCCEEDED(customSaveHr),
                  std::format(L"SomethingToSave failed after custom ViewerText settings. hr=0x{:08X}", static_cast<unsigned long>(customSaveHr)));
    if (SUCCEEDED(customSaveHr))
    {
        state.Require(somethingToSave == TRUE, L"Custom ViewerText configuration should be dirty.");
    }

    configuration             = nullptr;
    const HRESULT getCustomHr = info->GetConfiguration(&configuration);
    state.Require(SUCCEEDED(getCustomHr) && configuration != nullptr,
                  std::format(L"GetConfiguration failed after custom ViewerText settings. hr=0x{:08X}", static_cast<unsigned long>(getCustomHr)));
    if (configuration != nullptr)
    {
        const std::string_view configurationView(configuration);
        state.Require(configurationView.find("\"hexByteColorMode\":\"off\"") != std::string_view::npos,
                      L"ViewerText configuration JSON missing hexByteColorMode=off.");
        state.Require(configurationView.find("\"diffDefaultLayout\":\"inline\"") != std::string_view::npos,
                      L"ViewerText configuration JSON missing diffDefaultLayout=inline.");
        state.Require(configurationView.find("\"diffContextMode\":\"fullFileWhenAvailable\"") != std::string_view::npos,
                      L"ViewerText configuration JSON missing diffContextMode=fullFileWhenAvailable.");
        state.Require(configurationView.find("\"diffAutoOpenMode\":\"rawText\"") != std::string_view::npos,
                      L"ViewerText configuration JSON missing diffAutoOpenMode=rawText.");
    }

    const std::string resetConfiguration =
        R"json({"textBufferMiB":16,"hexBufferMiB":8,"showLineNumbers":"0","wrapText":"1","hexByteColorMode":"leadingNibble","diffDefaultLayout":"sideBySide","diffContextMode":"hunksOnly","diffAutoOpenMode":"parsed"})json";
    const HRESULT resetHr = info->SetConfiguration(resetConfiguration.c_str());
    state.Require(SUCCEEDED(resetHr),
                  std::format(L"SetConfiguration failed when resetting ViewerText settings. hr=0x{:08X}", static_cast<unsigned long>(resetHr)));
    if (FAILED(resetHr))
    {
        return false;
    }

    somethingToSave           = TRUE;
    const HRESULT resetSaveHr = info->SomethingToSave(&somethingToSave);
    state.Require(SUCCEEDED(resetSaveHr),
                  std::format(L"SomethingToSave failed after resetting ViewerText settings. hr=0x{:08X}", static_cast<unsigned long>(resetSaveHr)));
    if (SUCCEEDED(resetSaveHr))
    {
        state.Require(somethingToSave == FALSE, L"Reset ViewerText configuration should return to a clean state.");
    }

    configuration            = nullptr;
    const HRESULT getResetHr = info->GetConfiguration(&configuration);
    state.Require(SUCCEEDED(getResetHr) && configuration != nullptr,
                  std::format(L"GetConfiguration failed after resetting ViewerText settings. hr=0x{:08X}", static_cast<unsigned long>(getResetHr)));
    if (configuration != nullptr)
    {
        const std::string_view configurationView(configuration);
        state.Require(configurationView.find("\"hexByteColorMode\":\"leadingNibble\"") != std::string_view::npos,
                      L"ViewerText configuration JSON missing hexByteColorMode=leadingNibble after reset.");
        state.Require(configurationView.find("\"diffDefaultLayout\":\"sideBySide\"") != std::string_view::npos,
                      L"ViewerText configuration JSON missing diffDefaultLayout=sideBySide after reset.");
        state.Require(configurationView.find("\"diffContextMode\":\"hunksOnly\"") != std::string_view::npos,
                      L"ViewerText configuration JSON missing diffContextMode=hunksOnly after reset.");
        state.Require(configurationView.find("\"diffAutoOpenMode\":\"parsed\"") != std::string_view::npos,
                      L"ViewerText configuration JSON missing diffAutoOpenMode=parsed after reset.");
    }

    return state.failure.empty();
}

[[nodiscard]] bool TestViewerTextHexByteColorPerfScenario(CaseState& state) noexcept
{
    std::error_code ec;
    const std::filesystem::path tempRoot = std::filesystem::temp_directory_path(ec) / std::format(L"RedSalamander.ViewerTextPerf.{}", NewGuidText());
    state.Require(! ec, L"ViewerText perf scenario: failed to resolve the temp directory.");
    if (ec)
    {
        return false;
    }

    std::filesystem::create_directories(tempRoot, ec);
    state.Require(! ec, L"ViewerText perf scenario: failed to create the temp directory.");
    if (ec)
    {
        return false;
    }

    auto cleanupTempRoot = wil::scope_exit([&]() noexcept
    {
        std::error_code cleanupEc;
        std::filesystem::remove_all(tempRoot, cleanupEc);
    });

    static constexpr auto kFixtureBytes = std::to_array<std::byte>({
        std::byte{0x00}, std::byte{0x0F}, std::byte{0x10}, std::byte{0x1A}, std::byte{0x2B}, std::byte{0x3C}, std::byte{0x4D}, std::byte{0x5E},
        std::byte{0x6F}, std::byte{0x70}, std::byte{0x7F}, std::byte{0x80}, std::byte{0x8A}, std::byte{0x90}, std::byte{0xA5}, std::byte{0xBF},
        std::byte{0xC1}, std::byte{0xD2}, std::byte{0xE3}, std::byte{0xF4}, std::byte{0xFF}, std::byte{0x00}, std::byte{0x11}, std::byte{0x22},
        std::byte{0x33}, std::byte{0x44}, std::byte{0x55}, std::byte{0x66}, std::byte{0x77}, std::byte{0x88}, std::byte{0x99}, std::byte{0xAA},
    });

    const std::filesystem::path samplePath = tempRoot / L"viewertext-hex-byte-colors.bin";
    {
        wil::unique_handle file(CreateFileW(samplePath.c_str(), GENERIC_WRITE, 0u, nullptr, CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, nullptr));
        state.Require(file.is_valid(), L"ViewerText perf scenario: failed to create the binary fixture.");
        if (! file)
        {
            return false;
        }

        DWORD written      = 0u;
        const BOOL writeOk = WriteFile(file.get(), kFixtureBytes.data(), static_cast<DWORD>(kFixtureBytes.size()), &written, nullptr);
        state.Require(writeOk != FALSE && written == kFixtureBytes.size(), L"ViewerText perf scenario: failed to write the binary fixture.");
        if (writeOk == FALSE || written != kFixtureBytes.size())
        {
            return false;
        }
    }

    CreatedFileSystemInstance created{};
    const HRESULT fsHr = TryCreateFileSystemInstance(kBuiltinLocalFileSystemId, created);
    state.Require(SUCCEEDED(fsHr) && created.fileSystem,
                  std::format(L"ViewerText perf scenario: failed to create the local file system plugin. hr=0x{:08X}", static_cast<unsigned long>(fsHr)));
    if (FAILED(fsHr) || ! created.fileSystem)
    {
        return false;
    }

    const auto closeExistingViewerWindows = [&]() noexcept
    {
        for (;;)
        {
            HWND existing = FindWindowW(L"RedSalamander.ViewerText", nullptr);
            if (! existing)
            {
                break;
            }
            static_cast<void>(SendMessageW(existing, WM_CLOSE, 0, 0));
            static_cast<void>(WaitForWindowClosed(existing, SelfTest::Scale(std::chrono::milliseconds{5000})));
        }
    };

    const auto makeTheme = []() noexcept
    {
        const auto argbFromColorF = [](const D2D1::ColorF& color) noexcept
        {
            const auto clampByte = [](float value) noexcept -> uint32_t
            {
                value = std::clamp(value, 0.0f, 1.0f);
                return static_cast<uint32_t>(value * 255.0f + 0.5f);
            };

            const uint32_t a = clampByte(color.a);
            const uint32_t r = clampByte(color.r);
            const uint32_t g = clampByte(color.g);
            const uint32_t b = clampByte(color.b);
            return (a << 24) | (r << 16) | (g << 8) | b;
        };

        const AppTheme appTheme = ResolveAppTheme(ThemeMode::Light, L"viewer-text-diff-selftest");
        ViewerTheme theme{};
        theme.version                       = 4u;
        theme.dpi                           = 96u;
        theme.backgroundArgb                = argbFromColorF(appTheme.folderView.backgroundColor);
        theme.textArgb                      = argbFromColorF(appTheme.folderView.textNormal);
        theme.selectionBackgroundArgb       = argbFromColorF(appTheme.folderView.itemBackgroundSelected);
        theme.selectionTextArgb             = argbFromColorF(appTheme.folderView.textSelected);
        theme.accentArgb                    = argbFromColorF(appTheme.accent);
        theme.alertErrorBackgroundArgb      = argbFromColorF(appTheme.folderView.errorBackground);
        theme.alertErrorTextArgb            = argbFromColorF(appTheme.folderView.errorText);
        theme.alertWarningBackgroundArgb    = argbFromColorF(appTheme.folderView.warningBackground);
        theme.alertWarningTextArgb          = argbFromColorF(appTheme.folderView.warningText);
        theme.alertInfoBackgroundArgb       = argbFromColorF(appTheme.folderView.infoBackground);
        theme.alertInfoTextArgb             = argbFromColorF(appTheme.folderView.infoText);
        theme.darkMode                      = appTheme.dark ? TRUE : FALSE;
        theme.highContrast                  = appTheme.highContrast ? TRUE : FALSE;
        theme.rainbowMode                   = appTheme.menu.rainbowMode ? TRUE : FALSE;
        theme.darkBase                      = appTheme.menu.darkBase ? TRUE : FALSE;
        theme.diffAddedBackgroundArgb       = argbFromColorF(appTheme.viewerDiff.addedBackground);
        theme.diffRemovedBackgroundArgb     = argbFromColorF(appTheme.viewerDiff.removedBackground);
        theme.diffContextBackgroundArgb     = argbFromColorF(appTheme.viewerDiff.contextBackground);
        theme.diffHeaderBackgroundArgb      = argbFromColorF(appTheme.viewerDiff.headerBackground);
        theme.diffBannerBackgroundArgb      = argbFromColorF(appTheme.viewerDiff.bannerBackground);
        theme.diffPlaceholderBackgroundArgb = argbFromColorF(appTheme.viewerDiff.placeholderBackground);
        theme.diffDividerArgb               = argbFromColorF(appTheme.viewerDiff.divider);
        return theme;
    };

    const auto runScenario = [&](std::string_view configurationJson,
                                 std::wstring_view scenarioName,
                                 WndMsg::ViewerTextDebugHexByteColorMode expectedMode,
                                 bool expectColorized) noexcept
    {
        Trace(std::format(L"ViewerText hex byte color perf scenario: {}", scenarioName));
        closeExistingViewerWindows();

        Common::Settings::Settings isolatedSettings{};
        wil::com_ptr<IViewer> viewer;
        const HRESULT createHr = ViewerPluginManager::GetInstance().CreateViewerInstance(L"builtin/viewer-text", isolatedSettings, viewer);
        state.Require(
            SUCCEEDED(createHr) && viewer,
            std::format(L"ViewerText perf scenario '{}': failed to create the viewer. hr=0x{:08X}", scenarioName, static_cast<unsigned long>(createHr)));
        if (FAILED(createHr) || ! viewer)
        {
            return false;
        }

        wil::com_ptr<IInformations> info;
        const HRESULT infoHr = viewer->QueryInterface(__uuidof(IInformations), info.put_void());
        state.Require(
            SUCCEEDED(infoHr) && info,
            std::format(L"ViewerText perf scenario '{}': viewer missing IInformations. hr=0x{:08X}", scenarioName, static_cast<unsigned long>(infoHr)));
        if (FAILED(infoHr) || ! info)
        {
            return false;
        }

        const std::string configText(configurationJson);
        const HRESULT setHr = info->SetConfiguration(configText.c_str());
        state.Require(SUCCEEDED(setHr),
                      std::format(L"ViewerText perf scenario '{}': SetConfiguration failed. hr=0x{:08X}", scenarioName, static_cast<unsigned long>(setHr)));
        if (FAILED(setHr))
        {
            return false;
        }

        const ViewerTheme theme = makeTheme();
        const HRESULT themeHr   = viewer->SetTheme(&theme);
        state.Require(SUCCEEDED(themeHr),
                      std::format(L"ViewerText perf scenario '{}': SetTheme failed. hr=0x{:08X}", scenarioName, static_cast<unsigned long>(themeHr)));
        if (FAILED(themeHr))
        {
            return false;
        }

        const std::wstring samplePathText = samplePath.wstring();
        ViewerOpenContext context{};
        context.fileSystem     = created.fileSystem.get();
        context.fileSystemName = L"File System";
        context.focusedPath    = samplePathText.c_str();
        context.flags          = VIEWER_OPEN_FLAG_START_HEX;

        const HRESULT openHr = viewer->Open(&context);
        state.Require(SUCCEEDED(openHr),
                      std::format(L"ViewerText perf scenario '{}': Open failed. hr=0x{:08X}", scenarioName, static_cast<unsigned long>(openHr)));
        if (FAILED(openHr))
        {
            return false;
        }

        HWND viewerWindow = nullptr;
        auto closeViewer  = wil::scope_exit([&]() noexcept
        {
            static_cast<void>(viewer->Close());
            if (viewerWindow)
            {
                static_cast<void>(WaitForWindowClosed(viewerWindow, SelfTest::Scale(std::chrono::milliseconds{5000})));
            }
            closeExistingViewerWindows();
        });

        viewerWindow =
            WaitForWindow([]() noexcept { return FindWindowW(L"RedSalamander.ViewerText", nullptr); }, SelfTest::Scale(std::chrono::milliseconds{5000}));
        state.Require(viewerWindow != nullptr, std::format(L"ViewerText perf scenario '{}': viewer window did not appear.", scenarioName));
        if (! viewerWindow)
        {
            return false;
        }

        const auto tryGetSnapshot = [](HWND hwnd, WndMsg::ViewerTextDebugSnapshot& snapshot) noexcept
        {
            snapshot = {};
            return hwnd && IsWindow(hwnd) != FALSE && SendMessageW(hwnd, WndMsg::kViewerTextDebugGetSnapshot, 0, reinterpret_cast<LPARAM>(&snapshot)) != FALSE;
        };

        const auto waitForSnapshot = [&](auto&& predicate, WndMsg::ViewerTextDebugSnapshot& outSnapshot) noexcept
        {
            using namespace std::chrono_literals;

            WndMsg::ViewerTextDebugSnapshot snapshot{};
            const auto deadline = std::chrono::steady_clock::now() + SelfTest::Scale(std::chrono::milliseconds{5000});
            while (std::chrono::steady_clock::now() < deadline)
            {
                PumpPendingMessages();
                if (tryGetSnapshot(viewerWindow, snapshot) && predicate(snapshot))
                {
                    outSnapshot = snapshot;
                    return true;
                }
                std::this_thread::sleep_for(std::chrono::milliseconds{10});
            }

            if (tryGetSnapshot(viewerWindow, snapshot))
            {
                outSnapshot = snapshot;
            }
            return false;
        };

        WndMsg::ViewerTextDebugSnapshot initialSnapshot{};
        const bool snapshotReady = waitForSnapshot(
            [&](const WndMsg::ViewerTextDebugSnapshot& snapshot) noexcept
        {
            return snapshot.viewMode == WndMsg::ViewerTextDebugViewMode::Hex && snapshot.hexByteColorMode == expectedMode && snapshot.renderCount > 0u &&
                   snapshot.visibleByteCount > 0u &&
                   (expectColorized ? (snapshot.visibleColorizedByteCount > 0u && snapshot.visibleUniqueColorBucketCount > 1u)
                                    : snapshot.visibleColorizedByteCount == 0u);
        },
            initialSnapshot);
        state.Require(snapshotReady, std::format(L"ViewerText perf scenario '{}': initial hex snapshot did not reach the expected state.", scenarioName));
        if (! snapshotReady)
        {
            return false;
        }

        const HWND hexView = FindWindowExW(viewerWindow, nullptr, L"RedSalamander.ViewerText.HexView", nullptr);
        state.Require(hexView != nullptr, std::format(L"ViewerText perf scenario '{}': hex child window not found.", scenarioName));
        if (! hexView)
        {
            return false;
        }

        const uint64_t renderBeforeScroll = initialSnapshot.renderCount;
        static_cast<void>(SendMessageW(hexView, WM_VSCROLL, MAKEWPARAM(SB_LINEDOWN, 0), 0));

        WndMsg::ViewerTextDebugSnapshot scrolledSnapshot{};
        const bool scrolledReady = waitForSnapshot(
            [&](const WndMsg::ViewerTextDebugSnapshot& snapshot) noexcept
        {
            return snapshot.renderCount > renderBeforeScroll && snapshot.visibleByteCount > 0u &&
                   (expectColorized ? snapshot.visibleColorizedByteCount > 0u : snapshot.visibleColorizedByteCount == 0u);
        },
            scrolledSnapshot);
        state.Require(scrolledReady, std::format(L"ViewerText perf scenario '{}': scroll repaint snapshot did not arrive.", scenarioName));
        return scrolledReady;
    };

    const bool offOk = runScenario(R"json({"textBufferMiB":16,"hexBufferMiB":8,"showLineNumbers":"0","wrapText":"1","hexByteColorMode":"off"})json",
                                   L"off",
                                   WndMsg::ViewerTextDebugHexByteColorMode::Off,
                                   false);
    const bool leadingNibbleOk =
        runScenario(R"json({"textBufferMiB":16,"hexBufferMiB":8,"showLineNumbers":"0","wrapText":"1","hexByteColorMode":"leadingNibble"})json",
                    L"leadingNibble",
                    WndMsg::ViewerTextDebugHexByteColorMode::LeadingNibble,
                    true);

    return offOk && leadingNibbleOk && state.failure.empty();
}

[[nodiscard]] bool TestViewerTextDiffPerfScenario(CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    struct ScenarioMetrics
    {
        std::wstring name;
        std::filesystem::path diffPath;
        uint64_t diffBytes                             = 0u;
        uint64_t openToVisibleUs                       = 0u;
        uint64_t semanticRowPaintUs                    = 0u;
        uint64_t themeSwitchRepaintUs                  = 0u;
        uint64_t scrollRepaintUs                       = 0u;
        uint64_t hunkJumpToVisibleUs                   = 0u;
        uint64_t expandContextUs                       = 0u;
        uint64_t viewportRehydrateUs                   = 0u;
        uint64_t viewportBacktrackUs                   = 0u;
        size_t fileSectionCount                        = 0u;
        size_t visibleRowCount                         = 0u;
        size_t visibleStyledRowCount                   = 0u;
        size_t visibleContextRowCount                  = 0u;
        size_t visibleBannerRowCount                   = 0u;
        size_t visibleSplitRowCount                    = 0u;
        size_t placeholderRowCount                     = 0u;
        size_t placeholderBandCount                    = 0u;
        size_t deferredContextRowCount                 = 0u;
        size_t hydratedLogicalLineStart                = 0u;
        size_t hydratedLogicalLineEndExclusive         = 0u;
        size_t topVisibleLogicalLine                   = 0u;
        uint32_t sideBySideLeftPaneColumns             = 0u;
        uint32_t sideBySideRightPaneColumns            = 0u;
        uint32_t sideBySideSeparatorColumns            = 0u;
        uint64_t referencedBytesRead                   = 0u;
        uint64_t expandedReferencedBytesRead           = 0u;
        uint64_t viewportReferencedBytesRead           = 0u;
        uint64_t viewportReferencedBytesDelta          = 0u;
        uint64_t viewportBacktrackReferencedBytesDelta = 0u;
        bool diffExpandedContext                       = false;
        bool referencedFilesResolved                   = false;
        bool hasPlaceholderRows                        = false;
        bool paneLocalSideBySideLayout                 = false;
        bool viewportReferencedBytesIncreased          = false;
        bool viewportBacktrackReusedCachedBytes        = false;
    };

    constexpr WPARAM kViewerTextShowUnchangedCommand = 40213u;
    constexpr WPARAM kViewerTextNextHunkCommand      = 40214u;

    std::error_code ec;
    const std::filesystem::path tempRoot = std::filesystem::temp_directory_path(ec) / std::format(L"RedSalamander.ViewerTextDiffPerf.{}", NewGuidText());
    state.Require(! ec, L"ViewerText diff perf scenario: failed to resolve the temp directory.");
    if (ec)
    {
        return false;
    }

    std::filesystem::create_directories(tempRoot, ec);
    state.Require(! ec, L"ViewerText diff perf scenario: failed to create the temp directory.");
    if (ec)
    {
        return false;
    }

    auto cleanupTempRoot = wil::scope_exit([&]() noexcept
    {
        std::error_code cleanupEc;
        std::filesystem::remove_all(tempRoot, cleanupEc);
    });

    const auto makeStableFileBody = [](std::string_view stablePrefix,
                                       size_t lineCount,
                                       std::initializer_list<std::pair<size_t, std::string_view>> overrides,
                                       std::string_view stableSuffix = {}) -> std::string
    {
        std::string text;
        text.reserve(lineCount * (40u + stableSuffix.size()));

        for (size_t lineIndex = 1u; lineIndex <= lineCount; ++lineIndex)
        {
            std::string_view lineText;
            for (const auto& [overrideLine, overrideValue] : overrides)
            {
                if (overrideLine == lineIndex)
                {
                    lineText = overrideValue;
                    break;
                }
            }

            if (lineText.empty())
            {
                text += std::format("{} {:03}{}\n", stablePrefix, lineIndex, stableSuffix);
            }
            else
            {
                text += lineText;
                text.push_back('\n');
            }
        }

        return text;
    };

    const auto appendLargeMultiFileDiff = [](std::string& diffText, std::string_view oldPath, std::string_view newPath, int fileIndex, int hunkCount) noexcept
    {
        diffText += std::format("diff --git a/{} b/{}\n", oldPath, newPath);
        diffText += std::format("--- a/{}\n", oldPath);
        diffText += std::format("+++ b/{}\n", newPath);

        for (int hunkIndex = 0; hunkIndex < hunkCount; ++hunkIndex)
        {
            const int lineBase = 1 + (hunkIndex * 9);
            diffText += std::format("@@ -{},4 +{},4 @@\n", lineBase, lineBase);
            diffText += std::format(" shared file {:02} hunk {:02} line a\n", fileIndex, hunkIndex);
            diffText += std::format("-old file {:02} hunk {:02} line b\n", fileIndex, hunkIndex);
            diffText += std::format("+new file {:02} hunk {:02} line b\n", fileIndex, hunkIndex);
            diffText += std::format(" shared file {:02} hunk {:02} line c\n", fileIndex, hunkIndex);
            diffText += std::format(" shared file {:02} hunk {:02} line d\n", fileIndex, hunkIndex);
        }
    };

    const std::filesystem::path largeDiffPath    = tempRoot / L"viewertext-large-multi-file.diff";
    const std::filesystem::path resolvedOldPath  = tempRoot / L"viewertext-resolved-old.txt";
    const std::filesystem::path resolvedNewPath  = tempRoot / L"viewertext-resolved-new.txt";
    const std::filesystem::path resolvedDiffPath = tempRoot / L"viewertext-resolved-expand.diff";
    const std::filesystem::path missingDiffPath  = tempRoot / L"viewertext-missing-expand.diff";

    constexpr int kLargeFileCount = 12;
    constexpr int kLargeHunkCount = 16;
    constexpr int kResolvedLines  = 1024;

    std::string largeDiffText;
    largeDiffText.reserve(96u * 1024u);
    for (int fileIndex = 0; fileIndex < kLargeFileCount; ++fileIndex)
    {
        appendLargeMultiFileDiff(largeDiffText,
                                 std::format("viewertext-large-old-{:02}.txt", fileIndex),
                                 std::format("viewertext-large-new-{:02}.txt", fileIndex),
                                 fileIndex,
                                 kLargeHunkCount);
    }
    state.Require(SelfTest::WriteTextFile(largeDiffPath, largeDiffText), L"ViewerText diff perf scenario: failed to write the large multi-file diff fixture.");
    if (! state.failure.empty())
    {
        return false;
    }

    const std::string resolvedLineSuffix(192u, 'x');
    const std::string resolvedOldText = makeStableFileBody(
        "resolved stable line", kResolvedLines, {{20u, "resolved changed line 020 old"}, {920u, "resolved changed line 920 old"}}, resolvedLineSuffix);
    const std::string resolvedNewText = makeStableFileBody(
        "resolved stable line", kResolvedLines, {{20u, "resolved changed line 020 new"}, {920u, "resolved changed line 920 new"}}, resolvedLineSuffix);
    const uint64_t resolvedTotalReferencedBytes = static_cast<uint64_t>(resolvedOldText.size()) + static_cast<uint64_t>(resolvedNewText.size());
    state.Require(SelfTest::WriteTextFile(resolvedOldPath, resolvedOldText),
                  L"ViewerText diff perf scenario: failed to write the resolved old reference file.");
    state.Require(SelfTest::WriteTextFile(resolvedNewPath, resolvedNewText),
                  L"ViewerText diff perf scenario: failed to write the resolved new reference file.");
    if (! state.failure.empty())
    {
        return false;
    }

    const std::string resolvedDiffText = "diff --git a/viewertext-resolved-old.txt b/viewertext-resolved-new.txt\n"
                                         "--- a/viewertext-resolved-old.txt\n"
                                         "+++ b/viewertext-resolved-new.txt\n"
                                         "@@ -18,5 +18,5 @@\n"
                                         " resolved stable line 018\n"
                                         " resolved stable line 019\n"
                                         "-resolved changed line 020 old\n"
                                         "+resolved changed line 020 new\n"
                                         " resolved stable line 021\n"
                                         " resolved stable line 022\n"
                                         "@@ -918,5 +918,5 @@\n"
                                         " resolved stable line 918\n"
                                         " resolved stable line 919\n"
                                         "-resolved changed line 920 old\n"
                                         "+resolved changed line 920 new\n"
                                         " resolved stable line 921\n"
                                         " resolved stable line 922\n";
    state.Require(SelfTest::WriteTextFile(resolvedDiffPath, resolvedDiffText),
                  L"ViewerText diff perf scenario: failed to write the resolved expansion diff fixture.");
    if (! state.failure.empty())
    {
        return false;
    }

    const std::string missingDiffText = "diff --git a/viewertext-missing-left.txt b/viewertext-missing-right.txt\n"
                                        "--- a/viewertext-missing-left.txt\n"
                                        "+++ b/viewertext-missing-right.txt\n"
                                        "@@ -2 +2 @@\n"
                                        "-missing left beta\n"
                                        "+missing right beta\n"
                                        "@@ -7 +7 @@\n"
                                        "-missing left eta\n"
                                        "+missing right eta\n";
    state.Require(SelfTest::WriteTextFile(missingDiffPath, missingDiffText),
                  L"ViewerText diff perf scenario: failed to write the unresolved placeholder diff fixture.");
    if (! state.failure.empty())
    {
        return false;
    }

    const auto largeDiffBytes  = std::filesystem::file_size(largeDiffPath, ec);
    const bool largeDiffSizeOk = ! ec;
    ec.clear();
    const auto resolvedDiffBytes  = std::filesystem::file_size(resolvedDiffPath, ec);
    const bool resolvedDiffSizeOk = ! ec;
    ec.clear();
    const auto missingDiffBytes  = std::filesystem::file_size(missingDiffPath, ec);
    const bool missingDiffSizeOk = ! ec;
    ec.clear();

    state.Require(largeDiffSizeOk && resolvedDiffSizeOk && missingDiffSizeOk, L"ViewerText diff perf scenario: failed to query fixture sizes.");
    if (! state.failure.empty())
    {
        return false;
    }

    CreatedFileSystemInstance created{};
    const HRESULT fsHr = TryCreateFileSystemInstance(kBuiltinLocalFileSystemId, created);
    state.Require(SUCCEEDED(fsHr) && created.fileSystem,
                  std::format(L"ViewerText diff perf scenario: failed to create the local file system plugin. hr=0x{:08X}", static_cast<unsigned long>(fsHr)));
    if (FAILED(fsHr) || ! created.fileSystem)
    {
        return false;
    }

    const auto closeExistingViewerWindows = [&]() noexcept
    {
        for (;;)
        {
            HWND existing = FindWindowW(L"RedSalamander.ViewerText", nullptr);
            if (! existing)
            {
                break;
            }
            static_cast<void>(SendMessageW(existing, WM_CLOSE, 0, 0));
            static_cast<void>(WaitForWindowClosed(existing, SelfTest::Scale(5000ms)));
        }
    };

    const auto makeTheme = [](ThemeMode requestedMode, std::wstring_view rainbowSeed) noexcept
    {
        const auto argbFromColorF = [](const D2D1::ColorF& color) noexcept
        {
            const auto clampByte = [](float value) noexcept -> uint32_t
            {
                value = std::clamp(value, 0.0f, 1.0f);
                return static_cast<uint32_t>(value * 255.0f + 0.5f);
            };

            const uint32_t a = clampByte(color.a);
            const uint32_t r = clampByte(color.r);
            const uint32_t g = clampByte(color.g);
            const uint32_t b = clampByte(color.b);
            return (a << 24) | (r << 16) | (g << 8) | b;
        };

        const AppTheme appTheme = ResolveAppTheme(requestedMode, rainbowSeed);
        ViewerTheme theme{};
        theme.version                       = 4u;
        theme.dpi                           = 96u;
        theme.backgroundArgb                = argbFromColorF(appTheme.folderView.backgroundColor);
        theme.textArgb                      = argbFromColorF(appTheme.folderView.textNormal);
        theme.selectionBackgroundArgb       = argbFromColorF(appTheme.folderView.itemBackgroundSelected);
        theme.selectionTextArgb             = argbFromColorF(appTheme.folderView.textSelected);
        theme.accentArgb                    = argbFromColorF(appTheme.accent);
        theme.alertErrorBackgroundArgb      = argbFromColorF(appTheme.folderView.errorBackground);
        theme.alertErrorTextArgb            = argbFromColorF(appTheme.folderView.errorText);
        theme.alertWarningBackgroundArgb    = argbFromColorF(appTheme.folderView.warningBackground);
        theme.alertWarningTextArgb          = argbFromColorF(appTheme.folderView.warningText);
        theme.alertInfoBackgroundArgb       = argbFromColorF(appTheme.folderView.infoBackground);
        theme.alertInfoTextArgb             = argbFromColorF(appTheme.folderView.infoText);
        theme.darkMode                      = appTheme.dark ? TRUE : FALSE;
        theme.highContrast                  = appTheme.highContrast ? TRUE : FALSE;
        theme.rainbowMode                   = appTheme.menu.rainbowMode ? TRUE : FALSE;
        theme.darkBase                      = appTheme.menu.darkBase ? TRUE : FALSE;
        theme.diffAddedBackgroundArgb       = argbFromColorF(appTheme.viewerDiff.addedBackground);
        theme.diffRemovedBackgroundArgb     = argbFromColorF(appTheme.viewerDiff.removedBackground);
        theme.diffContextBackgroundArgb     = argbFromColorF(appTheme.viewerDiff.contextBackground);
        theme.diffHeaderBackgroundArgb      = argbFromColorF(appTheme.viewerDiff.headerBackground);
        theme.diffBannerBackgroundArgb      = argbFromColorF(appTheme.viewerDiff.bannerBackground);
        theme.diffPlaceholderBackgroundArgb = argbFromColorF(appTheme.viewerDiff.placeholderBackground);
        theme.diffDividerArgb               = argbFromColorF(appTheme.viewerDiff.divider);
        return theme;
    };

    const auto withOpenedDiffViewer =
        [&](std::wstring_view scenarioName, const std::filesystem::path& diffPath, std::string_view configurationJson, auto&& body) noexcept -> bool
    {
        Trace(std::format(L"ViewerText diff perf scenario: {}", scenarioName));
        closeExistingViewerWindows();

        Common::Settings::Settings isolatedSettings{};
        wil::com_ptr<IViewer> viewer;
        const HRESULT createHr = ViewerPluginManager::GetInstance().CreateViewerInstance(L"builtin/viewer-text", isolatedSettings, viewer);
        state.Require(
            SUCCEEDED(createHr) && viewer,
            std::format(L"ViewerText diff perf scenario '{}': failed to create the viewer. hr=0x{:08X}", scenarioName, static_cast<unsigned long>(createHr)));
        if (FAILED(createHr) || ! viewer)
        {
            return false;
        }

        wil::com_ptr<IInformations> info;
        const HRESULT infoHr = viewer->QueryInterface(__uuidof(IInformations), info.put_void());
        state.Require(
            SUCCEEDED(infoHr) && info,
            std::format(L"ViewerText diff perf scenario '{}': viewer missing IInformations. hr=0x{:08X}", scenarioName, static_cast<unsigned long>(infoHr)));
        if (FAILED(infoHr) || ! info)
        {
            return false;
        }

        const std::string configText(configurationJson);
        const HRESULT setHr = info->SetConfiguration(configText.c_str());
        state.Require(
            SUCCEEDED(setHr),
            std::format(L"ViewerText diff perf scenario '{}': SetConfiguration failed. hr=0x{:08X}", scenarioName, static_cast<unsigned long>(setHr)));
        if (FAILED(setHr))
        {
            return false;
        }

        const ViewerTheme theme = makeTheme(ThemeMode::Light, L"viewer-text-diff-perf-light");
        const HRESULT themeHr   = viewer->SetTheme(&theme);
        state.Require(SUCCEEDED(themeHr),
                      std::format(L"ViewerText diff perf scenario '{}': SetTheme failed. hr=0x{:08X}", scenarioName, static_cast<unsigned long>(themeHr)));
        if (FAILED(themeHr))
        {
            return false;
        }

        const std::wstring diffPathText = diffPath.wstring();
        ViewerOpenContext context{};
        context.fileSystem     = created.fileSystem.get();
        context.fileSystemName = L"File System";
        context.focusedPath    = diffPathText.c_str();

        const auto openStartedAt = std::chrono::steady_clock::now();
        const HRESULT openHr     = viewer->Open(&context);
        state.Require(SUCCEEDED(openHr),
                      std::format(L"ViewerText diff perf scenario '{}': Open failed. hr=0x{:08X}", scenarioName, static_cast<unsigned long>(openHr)));
        if (FAILED(openHr))
        {
            return false;
        }

        HWND viewerWindow = WaitForWindow([]() noexcept { return FindWindowW(L"RedSalamander.ViewerText", nullptr); }, SelfTest::Scale(5000ms));
        state.Require(viewerWindow != nullptr, std::format(L"ViewerText diff perf scenario '{}': viewer window did not appear.", scenarioName));
        if (! viewerWindow)
        {
            return false;
        }

        auto closeViewer = wil::scope_exit([&]() noexcept
        {
            static_cast<void>(viewer->Close());
            static_cast<void>(WaitForWindowClosed(viewerWindow, SelfTest::Scale(5000ms)));
            closeExistingViewerWindows();
        });

        const auto tryGetSnapshot = [](HWND hwnd, WndMsg::ViewerTextDebugSnapshot& snapshot) noexcept
        {
            snapshot = {};
            return hwnd && IsWindow(hwnd) != FALSE && SendMessageW(hwnd, WndMsg::kViewerTextDebugGetSnapshot, 0, reinterpret_cast<LPARAM>(&snapshot)) != FALSE;
        };

        const auto waitForSnapshot = [&](auto&& predicate, WndMsg::ViewerTextDebugSnapshot& outSnapshot, std::chrono::milliseconds timeout) noexcept
        {
            WndMsg::ViewerTextDebugSnapshot snapshot{};
            const auto deadline = std::chrono::steady_clock::now() + timeout;
            while (std::chrono::steady_clock::now() < deadline)
            {
                PumpPendingMessages();
                if (tryGetSnapshot(viewerWindow, snapshot) && predicate(snapshot))
                {
                    outSnapshot = snapshot;
                    return true;
                }
                std::this_thread::sleep_for(10ms);
            }

            if (tryGetSnapshot(viewerWindow, snapshot))
            {
                outSnapshot = snapshot;
            }
            return false;
        };

        return body(viewerWindow, viewer.get(), waitForSnapshot, openStartedAt);
    };

    ScenarioMetrics largeScenario{
        .name      = L"large_multi_file_side_by_side",
        .diffPath  = largeDiffPath,
        .diffBytes = largeDiffBytes,
    };
    const bool largeScenarioOk = withOpenedDiffViewer(
        largeScenario.name,
        largeDiffPath,
        R"json({"textBufferMiB":16,"hexBufferMiB":8,"showLineNumbers":"0","wrapText":"1","hexByteColorMode":"leadingNibble","diffDefaultLayout":"sideBySide","diffContextMode":"hunksOnly","diffAutoOpenMode":"parsed"})json",
        [&](HWND viewerWindow, IViewer* viewerInstance, auto&& waitForSnapshot, std::chrono::steady_clock::time_point openStartedAt) noexcept
    {
        WndMsg::ViewerTextDebugSnapshot initialSnapshot{};
        const bool snapshotReady = waitForSnapshot(
            [&](const WndMsg::ViewerTextDebugSnapshot& snapshot) noexcept
        {
            return snapshot.documentKind == WndMsg::ViewerTextDebugDocumentKind::Diff &&
                   snapshot.diffPresentation == WndMsg::ViewerTextDebugDiffPresentation::SideBySide && snapshot.diffParsedAvailable &&
                   snapshot.fileSectionCount == static_cast<size_t>(kLargeFileCount) && snapshot.renderCount > 0u && snapshot.visibleRowCount > 0u;
        },
            initialSnapshot,
            SelfTest::Scale(5000ms));
        state.Require(snapshotReady, L"ViewerText diff perf scenario 'large_multi_file_side_by_side': initial parsed diff snapshot did not arrive.");
        if (! snapshotReady)
        {
            return false;
        }

        largeScenario.openToVisibleUs            = Debug::Perf::ElapsedUs(openStartedAt);
        largeScenario.fileSectionCount           = initialSnapshot.fileSectionCount;
        largeScenario.visibleRowCount            = initialSnapshot.visibleRowCount;
        largeScenario.semanticRowPaintUs         = initialSnapshot.textLastPaintUs;
        largeScenario.visibleStyledRowCount      = initialSnapshot.visibleStyledRowCount;
        largeScenario.visibleContextRowCount     = initialSnapshot.visibleContextRowCount;
        largeScenario.visibleBannerRowCount      = initialSnapshot.visibleBannerRowCount;
        largeScenario.visibleSplitRowCount       = initialSnapshot.visibleSplitRowCount;
        largeScenario.topVisibleLogicalLine      = initialSnapshot.topVisibleLogicalLine;
        largeScenario.paneLocalSideBySideLayout  = initialSnapshot.paneLocalSideBySideLayout;
        largeScenario.sideBySideLeftPaneColumns  = initialSnapshot.sideBySideLeftPaneColumns;
        largeScenario.sideBySideRightPaneColumns = initialSnapshot.sideBySideRightPaneColumns;
        largeScenario.sideBySideSeparatorColumns = initialSnapshot.sideBySideSeparatorColumns;
        state.Require(initialSnapshot.visibleStyledRowCount > 0u,
                      L"ViewerText diff perf scenario 'large_multi_file_side_by_side': initial parsed diff paint should report visible styled rows.");
        state.Require(initialSnapshot.visibleContextRowCount > 0u,
                      L"ViewerText diff perf scenario 'large_multi_file_side_by_side': initial parsed diff paint should report visible context rows.");
        state.Require(initialSnapshot.paneLocalSideBySideLayout,
                      L"ViewerText diff perf scenario 'large_multi_file_side_by_side': pane-local side-by-side layout should be active.");
        state.Require(initialSnapshot.visibleSplitRowCount > 0u,
                      L"ViewerText diff perf scenario 'large_multi_file_side_by_side': initial parsed diff paint should report visible split rows.");
        state.Require(initialSnapshot.sideBySideLeftPaneColumns > 0u && initialSnapshot.sideBySideRightPaneColumns > 0u &&
                          initialSnapshot.sideBySideSeparatorColumns == 3u,
                      L"ViewerText diff perf scenario 'large_multi_file_side_by_side': pane-local side-by-side column metadata should be populated.");

        Debug::Perf::Emit(L"viewer.diff.open_to_first_visible_us",
                          largeScenario.name,
                          largeScenario.openToVisibleUs,
                          largeScenario.diffBytes,
                          largeScenario.fileSectionCount,
                          S_OK);
        Debug::Perf::Emit(L"viewer.diff.semantic_row_paint_us",
                          largeScenario.name,
                          largeScenario.semanticRowPaintUs,
                          largeScenario.visibleStyledRowCount,
                          largeScenario.visibleBannerRowCount,
                          S_OK);
        Debug::Perf::Emit(L"viewer.diff.visible_rows", largeScenario.name, 0u, largeScenario.visibleRowCount, largeScenario.fileSectionCount, S_OK);
        Debug::Perf::Emit(L"viewer.diff.visible_styled_rows", largeScenario.name, 0u, largeScenario.visibleStyledRowCount, largeScenario.visibleRowCount, S_OK);
        Debug::Perf::Emit(
            L"viewer.diff.visible_context_rows", largeScenario.name, 0u, largeScenario.visibleContextRowCount, largeScenario.visibleRowCount, S_OK);
        Debug::Perf::Emit(
            L"viewer.diff.visible_banner_rows", largeScenario.name, 0u, largeScenario.visibleBannerRowCount, largeScenario.visibleStyledRowCount, S_OK);
        Debug::Perf::Emit(L"viewer.diff.visible_split_rows", largeScenario.name, 0u, largeScenario.visibleSplitRowCount, largeScenario.visibleRowCount, S_OK);

        const ViewerTheme rainbowTheme = makeTheme(ThemeMode::Rainbow, L"viewer-text-diff-perf-rainbow");
        state.Require(rainbowTheme.rainbowMode != FALSE,
                      L"ViewerText diff perf scenario 'large_multi_file_side_by_side': rainbow perf theme should preserve ThemeMode::Rainbow.");
        const auto themeSwitchStartedAt = std::chrono::steady_clock::now();
        const HRESULT rainbowThemeHr    = viewerInstance ? viewerInstance->SetTheme(&rainbowTheme) : E_POINTER;
        state.Require(SUCCEEDED(rainbowThemeHr), L"ViewerText diff perf scenario 'large_multi_file_side_by_side': switching to the rainbow theme failed.");
        if (FAILED(rainbowThemeHr))
        {
            return false;
        }

        WndMsg::ViewerTextDebugSnapshot rainbowSnapshot{};
        const bool rainbowReady = waitForSnapshot(
            [&](const WndMsg::ViewerTextDebugSnapshot& snapshot) noexcept
        {
            return snapshot.renderCount > initialSnapshot.renderCount && snapshot.themeRainbow &&
                   snapshot.diffAddedBackgroundArgb == rainbowTheme.diffAddedBackgroundArgb &&
                   snapshot.diffRemovedBackgroundArgb == rainbowTheme.diffRemovedBackgroundArgb && snapshot.visibleStyledRowCount > 0u &&
                   snapshot.paneLocalSideBySideLayout && snapshot.visibleSplitRowCount > 0u;
        },
            rainbowSnapshot,
            SelfTest::Scale(5000ms));
        state.Require(rainbowReady, L"ViewerText diff perf scenario 'large_multi_file_side_by_side': rainbow theme repaint snapshot did not arrive.");
        if (! rainbowReady)
        {
            return false;
        }

        largeScenario.themeSwitchRepaintUs = Debug::Perf::ElapsedUs(themeSwitchStartedAt);
        Debug::Perf::Emit(L"viewer.diff.theme_switch_repaint_us",
                          largeScenario.name,
                          largeScenario.themeSwitchRepaintUs,
                          rainbowSnapshot.visibleStyledRowCount,
                          rainbowSnapshot.visibleBannerRowCount,
                          S_OK);

        const HWND textView = FindWindowExW(viewerWindow, nullptr, L"RedSalamander.ViewerText.TextView", nullptr);
        state.Require(textView != nullptr, L"ViewerText diff perf scenario 'large_multi_file_side_by_side': text child window not found.");
        if (! textView)
        {
            return false;
        }

        const auto scrollStartedAt = std::chrono::steady_clock::now();
        for (int step = 0; step < 3; ++step)
        {
            static_cast<void>(SendMessageW(textView, WM_VSCROLL, MAKEWPARAM(SB_LINEDOWN, 0), 0));
        }

        WndMsg::ViewerTextDebugSnapshot scrolledSnapshot{};
        const bool scrolledReady = waitForSnapshot([&](const WndMsg::ViewerTextDebugSnapshot& snapshot) noexcept {
            return snapshot.renderCount > initialSnapshot.renderCount && snapshot.topVisibleLogicalLine > initialSnapshot.topVisibleLogicalLine;
        }, scrolledSnapshot, SelfTest::Scale(5000ms));
        state.Require(scrolledReady, L"ViewerText diff perf scenario 'large_multi_file_side_by_side': scroll repaint snapshot did not arrive.");
        if (! scrolledReady)
        {
            return false;
        }

        largeScenario.scrollRepaintUs       = Debug::Perf::ElapsedUs(scrollStartedAt);
        largeScenario.topVisibleLogicalLine = scrolledSnapshot.topVisibleLogicalLine;

        Debug::Perf::Emit(L"viewer.diff.scroll_repaint_us",
                          largeScenario.name,
                          largeScenario.scrollRepaintUs,
                          largeScenario.topVisibleLogicalLine,
                          scrolledSnapshot.renderCount - initialSnapshot.renderCount,
                          S_OK);
        return true;
    });

    ScenarioMetrics resolvedScenario{
        .name      = L"resolved_expand_context",
        .diffPath  = resolvedDiffPath,
        .diffBytes = resolvedDiffBytes,
    };
    const bool resolvedScenarioOk = withOpenedDiffViewer(
        resolvedScenario.name,
        resolvedDiffPath,
        R"json({"textBufferMiB":16,"hexBufferMiB":8,"showLineNumbers":"0","wrapText":"1","hexByteColorMode":"leadingNibble","diffDefaultLayout":"sideBySide","diffContextMode":"hunksOnly","diffAutoOpenMode":"parsed"})json",
        [&](HWND viewerWindow, IViewer*, auto&& waitForSnapshot, std::chrono::steady_clock::time_point openStartedAt) noexcept
    {
        WndMsg::ViewerTextDebugSnapshot initialSnapshot{};
        const bool snapshotReady = waitForSnapshot(
            [](const WndMsg::ViewerTextDebugSnapshot& snapshot) noexcept
        {
            return snapshot.documentKind == WndMsg::ViewerTextDebugDocumentKind::Diff &&
                   snapshot.diffPresentation == WndMsg::ViewerTextDebugDiffPresentation::SideBySide && snapshot.diffParsedAvailable &&
                   ! snapshot.diffExpandedContext && snapshot.renderCount > 0u;
        },
            initialSnapshot,
            SelfTest::Scale(5000ms));
        state.Require(snapshotReady, L"ViewerText diff perf scenario 'resolved_expand_context': initial hunks-only snapshot did not arrive.");
        if (! snapshotReady)
        {
            return false;
        }

        resolvedScenario.openToVisibleUs  = Debug::Perf::ElapsedUs(openStartedAt);
        resolvedScenario.fileSectionCount = initialSnapshot.fileSectionCount;
        resolvedScenario.visibleRowCount  = initialSnapshot.visibleRowCount;
        state.Require(initialSnapshot.diffHunkCount > 1u,
                      L"ViewerText diff perf scenario 'resolved_expand_context': initial parsed diff should expose more than one navigable hunk.");

        Debug::Perf::Emit(L"viewer.diff.open_to_first_visible_us",
                          resolvedScenario.name,
                          resolvedScenario.openToVisibleUs,
                          resolvedScenario.diffBytes,
                          resolvedScenario.fileSectionCount,
                          S_OK);

        static_cast<void>(SendMessageW(viewerWindow, WndMsg::kViewerTextDebugSelectDiffHunk, 0u, 0));
        WndMsg::ViewerTextDebugSnapshot firstHunkSnapshot{};
        const bool firstHunkReady = waitForSnapshot(
            [&](const WndMsg::ViewerTextDebugSnapshot& snapshot) noexcept
        {
            return snapshot.renderCount > initialSnapshot.renderCount && snapshot.activeDiffHunkIndex == 0u &&
                   snapshot.topVisibleLogicalLine > initialSnapshot.topVisibleLogicalLine;
        },
            firstHunkSnapshot,
            SelfTest::Scale(5000ms));
        state.Require(firstHunkReady, L"ViewerText diff perf scenario 'resolved_expand_context': initial jump to the first hunk header did not arrive.");
        if (! firstHunkReady)
        {
            return false;
        }

        const auto hunkJumpStartedAt = std::chrono::steady_clock::now();
        static_cast<void>(SendMessageW(viewerWindow, WM_COMMAND, kViewerTextNextHunkCommand, 0));

        WndMsg::ViewerTextDebugSnapshot jumpedHunkSnapshot{};
        const bool jumpedHunkReady = waitForSnapshot(
            [&](const WndMsg::ViewerTextDebugSnapshot& snapshot) noexcept
        {
            return snapshot.renderCount > firstHunkSnapshot.renderCount && snapshot.activeDiffHunkIndex == 1u &&
                   snapshot.topVisibleLogicalLine > firstHunkSnapshot.topVisibleLogicalLine;
        },
            jumpedHunkSnapshot,
            SelfTest::Scale(5000ms));
        state.Require(jumpedHunkReady, L"ViewerText diff perf scenario 'resolved_expand_context': next-hunk jump snapshot did not arrive.");
        if (! jumpedHunkReady)
        {
            return false;
        }

        resolvedScenario.hunkJumpToVisibleUs = Debug::Perf::ElapsedUs(hunkJumpStartedAt);
        Debug::Perf::Emit(L"viewer.diff.hunk_jump_to_visible_us",
                          resolvedScenario.name,
                          resolvedScenario.hunkJumpToVisibleUs,
                          jumpedHunkSnapshot.diffHunkCount,
                          jumpedHunkSnapshot.activeDiffHunkIndex,
                          S_OK);

        static_cast<void>(SendMessageW(viewerWindow, WndMsg::kViewerTextDebugSelectDiffHunk, 0u, 0));
        WndMsg::ViewerTextDebugSnapshot restoredHunkSnapshot{};
        const bool restoredHunkReady = waitForSnapshot(
            [&](const WndMsg::ViewerTextDebugSnapshot& snapshot) noexcept
        {
            return snapshot.renderCount > jumpedHunkSnapshot.renderCount && snapshot.activeDiffHunkIndex == 0u &&
                   snapshot.topVisibleLogicalLine <= jumpedHunkSnapshot.topVisibleLogicalLine;
        },
            restoredHunkSnapshot,
            SelfTest::Scale(5000ms));
        state.Require(restoredHunkReady,
                      L"ViewerText diff perf scenario 'resolved_expand_context': debug reset to the first hunk did not arrive after hunk-jump timing.");
        if (! restoredHunkReady)
        {
            return false;
        }

        initialSnapshot = restoredHunkSnapshot;

        const bool hasClickableBanner = initialSnapshot.firstClickableBannerLogicalLine != static_cast<size_t>(-1);
        state.Require(hasClickableBanner,
                      L"ViewerText diff perf scenario 'resolved_expand_context': hunks-only parsed diff mode should expose a clickable hidden-context banner.");

        const auto expandStartedAt = std::chrono::steady_clock::now();
        if (hasClickableBanner)
        {
            const LRESULT clickResult = SendMessageW(
                viewerWindow, WndMsg::kViewerTextDebugClickTextLogicalLine, static_cast<WPARAM>(initialSnapshot.firstClickableBannerLogicalLine), 0);
            state.Require(clickResult != FALSE, L"ViewerText diff perf scenario 'resolved_expand_context': hidden-context banner click dispatch failed.");
            if (clickResult == FALSE)
            {
                return false;
            }
        }
        else
        {
            static_cast<void>(SendMessageW(viewerWindow, WM_COMMAND, kViewerTextShowUnchangedCommand, 0));
        }

        WndMsg::ViewerTextDebugSnapshot expandedSnapshot{};
        const bool expandedReady = waitForSnapshot(
            [&](const WndMsg::ViewerTextDebugSnapshot& snapshot) noexcept
        {
            return snapshot.documentKind == WndMsg::ViewerTextDebugDocumentKind::Diff &&
                   snapshot.diffPresentation == WndMsg::ViewerTextDebugDiffPresentation::SideBySide && snapshot.diffExpandedContext &&
                   snapshot.diffReferencedFilesResolved && snapshot.deferredContextRowCount > 0u && snapshot.renderCount > initialSnapshot.renderCount;
        },
            expandedSnapshot,
            SelfTest::Scale(5000ms));
        state.Require(expandedReady, L"ViewerText diff perf scenario 'resolved_expand_context': hidden-context banner reveal snapshot did not arrive.");
        if (! expandedReady)
        {
            return false;
        }

        resolvedScenario.expandContextUs                 = Debug::Perf::ElapsedUs(expandStartedAt);
        resolvedScenario.diffExpandedContext             = expandedSnapshot.diffExpandedContext;
        resolvedScenario.referencedFilesResolved         = expandedSnapshot.diffReferencedFilesResolved;
        resolvedScenario.deferredContextRowCount         = expandedSnapshot.deferredContextRowCount;
        resolvedScenario.hydratedLogicalLineStart        = expandedSnapshot.hydratedLogicalLineStart;
        resolvedScenario.hydratedLogicalLineEndExclusive = expandedSnapshot.hydratedLogicalLineEndExclusive;
        resolvedScenario.semanticRowPaintUs              = expandedSnapshot.textLastPaintUs;
        resolvedScenario.visibleStyledRowCount           = expandedSnapshot.visibleStyledRowCount;
        resolvedScenario.visibleContextRowCount          = expandedSnapshot.visibleContextRowCount;
        resolvedScenario.visibleBannerRowCount           = expandedSnapshot.visibleBannerRowCount;
        resolvedScenario.expandedReferencedBytesRead     = expandedSnapshot.referencedBytesRead;
        resolvedScenario.referencedBytesRead             = expandedSnapshot.referencedBytesRead;
        state.Require(
            expandedSnapshot.referencedBytesRead > 0u && expandedSnapshot.referencedBytesRead < resolvedTotalReferencedBytes,
            L"ViewerText diff perf scenario 'resolved_expand_context': unchanged-text expansion should read only a bounded prefix of large referenced files.");
        state.Require(expandedSnapshot.visibleStyledRowCount > 0u,
                      L"ViewerText diff perf scenario 'resolved_expand_context': expanded diff paint should report visible styled rows.");
        state.Require(expandedSnapshot.visibleContextRowCount > 0u,
                      L"ViewerText diff perf scenario 'resolved_expand_context': expanded diff paint should report visible context rows.");

        Debug::Perf::Emit(L"viewer.diff.expand_context_us",
                          resolvedScenario.name,
                          resolvedScenario.expandContextUs,
                          expandedSnapshot.visibleRowCount,
                          expandedSnapshot.fileSectionCount,
                          S_OK);
        Debug::Perf::Emit(L"viewer.diff.semantic_row_paint_us",
                          resolvedScenario.name,
                          resolvedScenario.semanticRowPaintUs,
                          resolvedScenario.visibleStyledRowCount,
                          resolvedScenario.visibleBannerRowCount,
                          S_OK);
        Debug::Perf::Emit(
            L"viewer.diff.visible_styled_rows", resolvedScenario.name, 0u, resolvedScenario.visibleStyledRowCount, expandedSnapshot.visibleRowCount, S_OK);
        Debug::Perf::Emit(
            L"viewer.diff.visible_context_rows", resolvedScenario.name, 0u, resolvedScenario.visibleContextRowCount, expandedSnapshot.visibleRowCount, S_OK);
        Debug::Perf::Emit(L"viewer.diff.visible_banner_rows",
                          resolvedScenario.name,
                          0u,
                          resolvedScenario.visibleBannerRowCount,
                          resolvedScenario.visibleStyledRowCount,
                          S_OK);
        Debug::Perf::Emit(
            L"viewer.diff.deferred_rows", resolvedScenario.name, 0u, resolvedScenario.deferredContextRowCount, expandedSnapshot.fileSectionCount, S_OK);
        Debug::Perf::Emit(
            L"viewer.diff.referenced_bytes_read", resolvedScenario.name, 0u, expandedSnapshot.referencedBytesRead, expandedSnapshot.fileSectionCount, S_OK);

        const HWND textView = FindWindowExW(viewerWindow, nullptr, L"RedSalamander.ViewerText.TextView", nullptr);
        state.Require(textView != nullptr, L"ViewerText diff perf scenario 'resolved_expand_context': text child window not found for viewport rehydration.");
        if (! textView)
        {
            return false;
        }

        constexpr int kMaxViewportPageDownSteps = 24;
        const auto viewportRehydrateStartedAt   = std::chrono::steady_clock::now();
        WndMsg::ViewerTextDebugSnapshot viewportSnapshot{};
        WndMsg::ViewerTextDebugSnapshot latestViewportSnapshot = expandedSnapshot;
        bool viewportReady                                     = false;
        bool viewportReadGrowthObserved                        = false;
        for (int step = 0; step < kMaxViewportPageDownSteps; ++step)
        {
            const uint64_t previousRenderCount         = latestViewportSnapshot.renderCount;
            const size_t previousTopVisibleLogicalLine = latestViewportSnapshot.topVisibleLogicalLine;

            static_cast<void>(SendMessageW(textView, WM_VSCROLL, MAKEWPARAM(SB_PAGEDOWN, 0), 0));

            WndMsg::ViewerTextDebugSnapshot candidateSnapshot{};
            const bool candidateReady = waitForSnapshot([&](const WndMsg::ViewerTextDebugSnapshot& snapshot) noexcept {
                return snapshot.renderCount > previousRenderCount && snapshot.topVisibleLogicalLine > previousTopVisibleLogicalLine;
            }, candidateSnapshot, SelfTest::Scale(5000ms));
            state.Require(candidateReady, L"ViewerText diff perf scenario 'resolved_expand_context': viewport rehydration snapshot did not arrive.");
            if (! candidateReady)
            {
                return false;
            }

            latestViewportSnapshot = candidateSnapshot;
            viewportReady          = true;
            if (candidateSnapshot.referencedBytesRead > expandedSnapshot.referencedBytesRead)
            {
                viewportSnapshot           = candidateSnapshot;
                viewportReadGrowthObserved = true;
                break;
            }
        }

        if (! viewportReady)
        {
            state.Require(false, L"ViewerText diff perf scenario 'resolved_expand_context': viewport rehydration did not advance after page-down input.");
            return false;
        }
        if (! viewportReadGrowthObserved)
        {
            viewportSnapshot = latestViewportSnapshot;
        }

        resolvedScenario.viewportRehydrateUs              = Debug::Perf::ElapsedUs(viewportRehydrateStartedAt);
        resolvedScenario.topVisibleLogicalLine            = viewportSnapshot.topVisibleLogicalLine;
        resolvedScenario.hydratedLogicalLineStart         = viewportSnapshot.hydratedLogicalLineStart;
        resolvedScenario.hydratedLogicalLineEndExclusive  = viewportSnapshot.hydratedLogicalLineEndExclusive;
        resolvedScenario.viewportReferencedBytesRead      = viewportSnapshot.referencedBytesRead;
        resolvedScenario.viewportReferencedBytesDelta     = viewportSnapshot.referencedBytesRead > expandedSnapshot.referencedBytesRead
                                                                ? (viewportSnapshot.referencedBytesRead - expandedSnapshot.referencedBytesRead)
                                                                : 0u;
        resolvedScenario.referencedBytesRead              = viewportSnapshot.referencedBytesRead;
        resolvedScenario.viewportReferencedBytesIncreased = viewportReadGrowthObserved;
        state.Require(viewportSnapshot.hydratedLogicalLineStart > expandedSnapshot.hydratedLogicalLineStart,
                      L"ViewerText diff perf scenario 'resolved_expand_context': viewport rehydration should move the hydrated logical window.");
        state.Require(viewportReadGrowthObserved,
                      L"ViewerText diff perf scenario 'resolved_expand_context': viewport rehydration should trigger additional bounded referenced-file reads "
                      L"once the viewport advances beyond the initial hydrated prefix.");
        state.Require(
            viewportSnapshot.referencedBytesRead > 0u && viewportSnapshot.referencedBytesRead < resolvedTotalReferencedBytes,
            L"ViewerText diff perf scenario 'resolved_expand_context': viewport rehydration should keep referenced-file reads bounded below full-file size.");

        Debug::Perf::Emit(L"viewer.diff.viewport_rehydrate_us",
                          resolvedScenario.name,
                          resolvedScenario.viewportRehydrateUs,
                          viewportSnapshot.visibleRowCount,
                          viewportSnapshot.fileSectionCount,
                          S_OK);
        Debug::Perf::Emit(L"viewer.diff.viewport_referenced_bytes_read",
                          resolvedScenario.name,
                          0u,
                          resolvedScenario.viewportReferencedBytesRead,
                          viewportSnapshot.fileSectionCount,
                          S_OK);
        Debug::Perf::Emit(L"viewer.diff.viewport_referenced_bytes_delta",
                          resolvedScenario.name,
                          0u,
                          resolvedScenario.viewportReferencedBytesDelta,
                          viewportSnapshot.fileSectionCount,
                          S_OK);

        const auto viewportBacktrackStartedAt       = std::chrono::steady_clock::now();
        constexpr int kViewportBacktrackPageUpSteps = 8;
        for (int step = 0; step < kViewportBacktrackPageUpSteps; ++step)
        {
            static_cast<void>(SendMessageW(textView, WM_VSCROLL, MAKEWPARAM(SB_PAGEUP, 0), 0));
        }

        WndMsg::ViewerTextDebugSnapshot backtrackSnapshot{};
        const bool backtrackReady = waitForSnapshot([&](const WndMsg::ViewerTextDebugSnapshot& snapshot) noexcept {
            return snapshot.renderCount > viewportSnapshot.renderCount && snapshot.topVisibleLogicalLine < viewportSnapshot.topVisibleLogicalLine;
        }, backtrackSnapshot, SelfTest::Scale(5000ms));
        state.Require(backtrackReady, L"ViewerText diff perf scenario 'resolved_expand_context': viewport backtrack snapshot did not arrive.");
        if (! backtrackReady)
        {
            return false;
        }

        resolvedScenario.viewportBacktrackUs                   = Debug::Perf::ElapsedUs(viewportBacktrackStartedAt);
        resolvedScenario.viewportBacktrackReferencedBytesDelta = backtrackSnapshot.referencedBytesRead > viewportSnapshot.referencedBytesRead
                                                                     ? (backtrackSnapshot.referencedBytesRead - viewportSnapshot.referencedBytesRead)
                                                                     : 0u;
        resolvedScenario.viewportBacktrackReusedCachedBytes    = (backtrackSnapshot.referencedBytesRead == viewportSnapshot.referencedBytesRead);
        state.Require(resolvedScenario.viewportBacktrackReusedCachedBytes,
                      L"ViewerText diff perf scenario 'resolved_expand_context': scrolling back to an already hydrated viewport range should not reread "
                      L"referenced-file bytes.");

        Debug::Perf::Emit(L"viewer.diff.viewport_backtrack_us",
                          resolvedScenario.name,
                          resolvedScenario.viewportBacktrackUs,
                          backtrackSnapshot.visibleRowCount,
                          backtrackSnapshot.fileSectionCount,
                          S_OK);
        Debug::Perf::Emit(L"viewer.diff.viewport_backtrack_referenced_bytes_delta",
                          resolvedScenario.name,
                          0u,
                          resolvedScenario.viewportBacktrackReferencedBytesDelta,
                          backtrackSnapshot.fileSectionCount,
                          S_OK);
        return true;
    });

    ScenarioMetrics missingScenario{
        .name      = L"unresolved_placeholder",
        .diffPath  = missingDiffPath,
        .diffBytes = missingDiffBytes,
    };
    const bool missingScenarioOk = withOpenedDiffViewer(
        missingScenario.name,
        missingDiffPath,
        R"json({"textBufferMiB":16,"hexBufferMiB":8,"showLineNumbers":"0","wrapText":"1","hexByteColorMode":"leadingNibble","diffDefaultLayout":"sideBySide","diffContextMode":"fullFileWhenAvailable","diffAutoOpenMode":"parsed"})json",
        [&](HWND, IViewer*, auto&& waitForSnapshot, std::chrono::steady_clock::time_point openStartedAt) noexcept
    {
        WndMsg::ViewerTextDebugSnapshot missingSnapshot{};
        const bool snapshotReady = waitForSnapshot(
            [](const WndMsg::ViewerTextDebugSnapshot& snapshot) noexcept
        {
            return snapshot.documentKind == WndMsg::ViewerTextDebugDocumentKind::Diff &&
                   snapshot.diffPresentation == WndMsg::ViewerTextDebugDiffPresentation::SideBySide && snapshot.diffParsedAvailable &&
                   snapshot.diffHasPlaceholderRows && snapshot.placeholderRowCount > 0u && snapshot.placeholderBandCount > 0u &&
                   ! snapshot.diffReferencedFilesResolved && snapshot.renderCount > 0u;
        },
            missingSnapshot,
            SelfTest::Scale(5000ms));
        state.Require(snapshotReady, L"ViewerText diff perf scenario 'unresolved_placeholder': placeholder snapshot did not arrive.");
        if (! snapshotReady)
        {
            return false;
        }

        missingScenario.openToVisibleUs         = Debug::Perf::ElapsedUs(openStartedAt);
        missingScenario.fileSectionCount        = missingSnapshot.fileSectionCount;
        missingScenario.visibleRowCount         = missingSnapshot.visibleRowCount;
        missingScenario.semanticRowPaintUs      = missingSnapshot.textLastPaintUs;
        missingScenario.visibleStyledRowCount   = missingSnapshot.visibleStyledRowCount;
        missingScenario.visibleContextRowCount  = missingSnapshot.visibleContextRowCount;
        missingScenario.visibleBannerRowCount   = missingSnapshot.visibleBannerRowCount;
        missingScenario.placeholderRowCount     = missingSnapshot.placeholderRowCount;
        missingScenario.placeholderBandCount    = missingSnapshot.placeholderBandCount;
        missingScenario.hasPlaceholderRows      = missingSnapshot.diffHasPlaceholderRows;
        missingScenario.referencedFilesResolved = missingSnapshot.diffReferencedFilesResolved;
        state.Require(missingSnapshot.visibleStyledRowCount > 0u,
                      L"ViewerText diff perf scenario 'unresolved_placeholder': placeholder diff paint should report visible styled rows.");

        Debug::Perf::Emit(L"viewer.diff.open_to_first_visible_us",
                          missingScenario.name,
                          missingScenario.openToVisibleUs,
                          missingScenario.diffBytes,
                          missingScenario.fileSectionCount,
                          S_OK);
        Debug::Perf::Emit(L"viewer.diff.semantic_row_paint_us",
                          missingScenario.name,
                          missingScenario.semanticRowPaintUs,
                          missingScenario.visibleStyledRowCount,
                          missingScenario.visibleBannerRowCount,
                          S_OK);
        Debug::Perf::Emit(
            L"viewer.diff.visible_styled_rows", missingScenario.name, 0u, missingScenario.visibleStyledRowCount, missingScenario.visibleRowCount, S_OK);
        Debug::Perf::Emit(
            L"viewer.diff.visible_context_rows", missingScenario.name, 0u, missingScenario.visibleContextRowCount, missingScenario.visibleRowCount, S_OK);
        Debug::Perf::Emit(
            L"viewer.diff.visible_banner_rows", missingScenario.name, 0u, missingScenario.visibleBannerRowCount, missingScenario.visibleStyledRowCount, S_OK);
        Debug::Perf::Emit(
            L"viewer.diff.placeholder_rows", missingScenario.name, 0u, missingScenario.placeholderRowCount, missingScenario.fileSectionCount, S_OK);
        Debug::Perf::Emit(
            L"viewer.diff.placeholder_bands", missingScenario.name, 0u, missingScenario.placeholderBandCount, missingScenario.fileSectionCount, S_OK);
        return true;
    });

    const bool scenariosOk = largeScenarioOk && resolvedScenarioOk && missingScenarioOk && state.failure.empty();
    if (! scenariosOk)
    {
        return false;
    }

    const std::wstring diffPerfArtifactText          = std::format(L"{{\n"
                                                                   L"  \"case\": \"viewer_text_diff_perf\",\n"
                                                                   L"  \"tempRoot\": \"{}\",\n"
                                                                   L"  \"scenarios\": [\n"
                                                                   L"    {{\n"
                                                                   L"      \"name\": \"{}\",\n"
                                                                   L"      \"diffPath\": \"{}\",\n"
                                                                   L"      \"diffBytes\": {},\n"
                                                                   L"      \"fileSectionCount\": {},\n"
                                                                   L"      \"visibleRowCount\": {},\n"
                                                                   L"      \"visibleStyledRowCount\": {},\n"
                                                                   L"      \"visibleContextRowCount\": {},\n"
                                                                   L"      \"visibleBannerRowCount\": {},\n"
                                                                   L"      \"visibleSplitRowCount\": {},\n"
                                                                   L"      \"paneLocalSideBySideLayout\": {},\n"
                                                                   L"      \"sideBySideLeftPaneColumns\": {},\n"
                                                                   L"      \"sideBySideRightPaneColumns\": {},\n"
                                                                   L"      \"sideBySideSeparatorColumns\": {},\n"
                                                                   L"      \"openToFirstVisibleUs\": {},\n"
                                                                   L"      \"semanticRowPaintUs\": {},\n"
                                                                   L"      \"themeSwitchRepaintUs\": {},\n"
                                                                   L"      \"scrollRepaintUs\": {},\n"
                                                                   L"      \"topVisibleLogicalLine\": {}\n"
                                                                   L"    }},\n"
                                                                   L"    {{\n"
                                                                   L"      \"name\": \"{}\",\n"
                                                                   L"      \"diffPath\": \"{}\",\n"
                                                                   L"      \"diffBytes\": {},\n"
                                                                   L"      \"fileSectionCount\": {},\n"
                                                                   L"      \"visibleRowCount\": {},\n"
                                                                   L"      \"visibleStyledRowCount\": {},\n"
                                                                   L"      \"visibleContextRowCount\": {},\n"
                                                                   L"      \"visibleBannerRowCount\": {},\n"
                                                                   L"      \"openToFirstVisibleUs\": {},\n"
                                                                   L"      \"hunkJumpToVisibleUs\": {},\n"
                                                                   L"      \"semanticRowPaintUs\": {},\n"
                                                                   L"      \"expandContextUs\": {},\n"
                                                                   L"      \"viewportRehydrateUs\": {},\n"
                                                                   L"      \"viewportBacktrackUs\": {},\n"
                                                                   L"      \"deferredContextRowCount\": {},\n"
                                                                   L"      \"hydratedLogicalLineStart\": {},\n"
                                                                   L"      \"hydratedLogicalLineEndExclusive\": {},\n"
                                                                   L"      \"expandedReferencedBytesRead\": {},\n"
                                                                   L"      \"viewportReferencedBytesRead\": {},\n"
                                                                   L"      \"viewportReferencedBytesDelta\": {},\n"
                                                                   L"      \"viewportBacktrackReferencedBytesDelta\": {},\n"
                                                                   L"      \"viewportReferencedBytesIncreased\": {},\n"
                                                                   L"      \"viewportBacktrackReusedCachedBytes\": {},\n"
                                                                   L"      \"referencedBytesRead\": {},\n"
                                                                   L"      \"diffExpandedContext\": {},\n"
                                                                   L"      \"referencedFilesResolved\": {}\n"
                                                                   L"    }},\n"
                                                                   L"    {{\n"
                                                                   L"      \"name\": \"{}\",\n"
                                                                   L"      \"diffPath\": \"{}\",\n"
                                                                   L"      \"diffBytes\": {},\n"
                                                                   L"      \"fileSectionCount\": {},\n"
                                                                   L"      \"visibleRowCount\": {},\n"
                                                                   L"      \"visibleStyledRowCount\": {},\n"
                                                                   L"      \"visibleContextRowCount\": {},\n"
                                                                   L"      \"visibleBannerRowCount\": {},\n"
                                                                   L"      \"openToFirstVisibleUs\": {},\n"
                                                                   L"      \"semanticRowPaintUs\": {},\n"
                                                                   L"      \"placeholderRowCount\": {},\n"
                                                                   L"      \"placeholderBandCount\": {},\n"
                                                                   L"      \"hasPlaceholderRows\": {},\n"
                                                                   L"      \"referencedFilesResolved\": {}\n"
                                                                   L"    }}\n"
                                                                   L"  ]\n"
                                                                   L"}}\n",
                                                                   tempRoot.generic_wstring(),
                                                                   largeScenario.name,
                                                                   largeScenario.diffPath.generic_wstring(),
                                                                   largeScenario.diffBytes,
                                                                   largeScenario.fileSectionCount,
                                                                   largeScenario.visibleRowCount,
                                                                   largeScenario.visibleStyledRowCount,
                                                                   largeScenario.visibleContextRowCount,
                                                                   largeScenario.visibleBannerRowCount,
                                                                   largeScenario.visibleSplitRowCount,
                                                                   largeScenario.paneLocalSideBySideLayout ? L"true" : L"false",
                                                                   largeScenario.sideBySideLeftPaneColumns,
                                                                   largeScenario.sideBySideRightPaneColumns,
                                                                   largeScenario.sideBySideSeparatorColumns,
                                                                   largeScenario.openToVisibleUs,
                                                                   largeScenario.semanticRowPaintUs,
                                                                   largeScenario.themeSwitchRepaintUs,
                                                                   largeScenario.scrollRepaintUs,
                                                                   largeScenario.topVisibleLogicalLine,
                                                                   resolvedScenario.name,
                                                                   resolvedScenario.diffPath.generic_wstring(),
                                                                   resolvedScenario.diffBytes,
                                                                   resolvedScenario.fileSectionCount,
                                                                   resolvedScenario.visibleRowCount,
                                                                   resolvedScenario.visibleStyledRowCount,
                                                                   resolvedScenario.visibleContextRowCount,
                                                                   resolvedScenario.visibleBannerRowCount,
                                                                   resolvedScenario.openToVisibleUs,
                                                                   resolvedScenario.hunkJumpToVisibleUs,
                                                                   resolvedScenario.semanticRowPaintUs,
                                                                   resolvedScenario.expandContextUs,
                                                                   resolvedScenario.viewportRehydrateUs,
                                                                   resolvedScenario.viewportBacktrackUs,
                                                                   resolvedScenario.deferredContextRowCount,
                                                                   resolvedScenario.hydratedLogicalLineStart,
                                                                   resolvedScenario.hydratedLogicalLineEndExclusive,
                                                                   resolvedScenario.expandedReferencedBytesRead,
                                                                   resolvedScenario.viewportReferencedBytesRead,
                                                                   resolvedScenario.viewportReferencedBytesDelta,
                                                                   resolvedScenario.viewportBacktrackReferencedBytesDelta,
                                                                   resolvedScenario.viewportReferencedBytesIncreased ? L"true" : L"false",
                                                                   resolvedScenario.viewportBacktrackReusedCachedBytes ? L"true" : L"false",
                                                                   resolvedScenario.referencedBytesRead,
                                                                   resolvedScenario.diffExpandedContext ? L"true" : L"false",
                                                                   resolvedScenario.referencedFilesResolved ? L"true" : L"false",
                                                                   missingScenario.name,
                                                                   missingScenario.diffPath.generic_wstring(),
                                                                   missingScenario.diffBytes,
                                                                   missingScenario.fileSectionCount,
                                                                   missingScenario.visibleRowCount,
                                                                   missingScenario.visibleStyledRowCount,
                                                                   missingScenario.visibleContextRowCount,
                                                                   missingScenario.visibleBannerRowCount,
                                                                   missingScenario.openToVisibleUs,
                                                                   missingScenario.semanticRowPaintUs,
                                                                   missingScenario.placeholderRowCount,
                                                                   missingScenario.placeholderBandCount,
                                                                   missingScenario.hasPlaceholderRows ? L"true" : L"false",
                                                                   missingScenario.referencedFilesResolved ? L"true" : L"false");
    const std::filesystem::path diffPerfArtifactPath = SelfTest::GetPerfArtifactPath(L"viewer_text_diff_perf_metrics.json");
    const bool diffPerfArtifactWriteOk               = ! diffPerfArtifactPath.empty() && SelfTest::WriteTextFile(diffPerfArtifactPath, diffPerfArtifactText);
    const bool diffPerfArtifactExists                = ! diffPerfArtifactPath.empty() && SelfTest::PathExists(diffPerfArtifactPath);
    SelfTest::AppendSuiteTrace(SelfTest::SelfTestSuite::Commands,
                               std::format(L"viewer_text_diff_perf artifact path='{}' writeOk={} existsAfterWrite={}",
                                           diffPerfArtifactPath.native(),
                                           diffPerfArtifactWriteOk,
                                           diffPerfArtifactExists));
    state.Require(diffPerfArtifactWriteOk && diffPerfArtifactExists, L"Failed to write ViewerText diff perf metrics artifact.");

    return state.failure.empty();
}

[[nodiscard]] bool RunStandaloneViewerCloseRoundTrip(CaseState& state,
                                                     std::wstring_view scenarioName,
                                                     std::wstring_view pluginId,
                                                     std::wstring_view windowClassName,
                                                     const std::filesystem::path& samplePath,
                                                     const DWORD openFlags = 0u) noexcept
{
    using namespace std::chrono_literals;

    CreatedFileSystemInstance created{};
    const HRESULT fsHr = TryCreateFileSystemInstance(kBuiltinLocalFileSystemId, created);
    state.Require(SUCCEEDED(fsHr) && created.fileSystem,
                  std::format(L"{}: failed to create the local file system plugin. hr=0x{:08X}", scenarioName, static_cast<unsigned long>(fsHr)));
    if (FAILED(fsHr) || ! created.fileSystem)
    {
        return false;
    }

    const std::wstring windowClass(windowClassName);
    const auto closeExistingViewerWindows = [&]() noexcept
    {
        for (;;)
        {
            HWND existing = FindWindowW(windowClass.c_str(), nullptr);
            if (! existing)
            {
                break;
            }
            static_cast<void>(SendMessageW(existing, WM_CLOSE, 0, 0));
            static_cast<void>(WaitForWindowClosed(existing, SelfTest::Scale(5000ms)));
        }
    };

    closeExistingViewerWindows();

    Common::Settings::Settings isolatedSettings{};
    wil::com_ptr<IViewer> viewer;
    const HRESULT createHr = ViewerPluginManager::GetInstance().CreateViewerInstance(pluginId, isolatedSettings, viewer);
    state.Require(SUCCEEDED(createHr) && viewer,
                  std::format(L"{}: failed to create viewer '{}' hr=0x{:08X}", scenarioName, pluginId, static_cast<unsigned long>(createHr)));
    if (FAILED(createHr) || ! viewer)
    {
        return false;
    }

    const std::wstring samplePathText = samplePath.wstring();
    ViewerOpenContext context{};
    context.fileSystem     = created.fileSystem.get();
    context.fileSystemName = L"File System";
    context.focusedPath    = samplePathText.c_str();
    context.flags          = static_cast<ViewerOpenFlags>(openFlags);

    HWND viewerWindow   = nullptr;
    bool closeRequested = false;
    auto cleanup        = wil::scope_exit([&]() noexcept
    {
        if (viewer && ! closeRequested)
        {
            static_cast<void>(viewer->Close());
        }
        if (viewerWindow)
        {
            static_cast<void>(WaitForWindowClosed(viewerWindow, SelfTest::Scale(5000ms)));
        }
        closeExistingViewerWindows();
    });

    const auto openStartedAt = std::chrono::steady_clock::now();
    const HRESULT openHr     = viewer->Open(&context);
    state.Require(SUCCEEDED(openHr), std::format(L"{}: Open failed for '{}' hr=0x{:08X}", scenarioName, samplePathText, static_cast<unsigned long>(openHr)));
    if (FAILED(openHr))
    {
        return false;
    }

    viewerWindow = WaitForWindow(
        [&]() noexcept
    {
        const HWND hwnd = FindWindowW(windowClass.c_str(), nullptr);
        return (hwnd && IsWindowVisible(hwnd) != FALSE) ? hwnd : nullptr;
    },
        SelfTest::Scale(5000ms));
    state.Require(viewerWindow != nullptr, std::format(L"{}: viewer window '{}' did not appear.", scenarioName, windowClassName));
    if (! viewerWindow)
    {
        return false;
    }

    PumpPendingMessages();
    std::this_thread::sleep_for(50ms);
    PumpPendingMessages();

    const auto closeStartedAt = std::chrono::steady_clock::now();
    const HRESULT closeHr     = viewer->Close();
    closeRequested            = true;
    state.Require(SUCCEEDED(closeHr), std::format(L"{}: Close failed. hr=0x{:08X}", scenarioName, static_cast<unsigned long>(closeHr)));
    if (FAILED(closeHr))
    {
        return false;
    }

    const bool closed = WaitForWindowClosed(viewerWindow, SelfTest::Scale(5000ms));
    state.Require(closed, std::format(L"{}: viewer window did not close.", scenarioName));
    if (! closed)
    {
        return false;
    }

    const auto openToVisibleUs = std::chrono::duration_cast<std::chrono::microseconds>(closeStartedAt - openStartedAt).count();
    const auto closeToClosedUs = std::chrono::duration_cast<std::chrono::microseconds>(std::chrono::steady_clock::now() - closeStartedAt).count();
    SelfTest::AppendSuiteTrace(
        SelfTest::SelfTestSuite::Commands,
        std::format(
            L"{} pluginId='{}' path='{}' openToVisibleUs={} closeToClosedUs={}", scenarioName, pluginId, samplePathText, openToVisibleUs, closeToClosedUs));

    viewerWindow = nullptr;
    return state.failure.empty();
}

[[nodiscard]] bool TestViewerPECloseRoundTrip(CaseState& state) noexcept
{
    wchar_t modulePath[MAX_PATH] = {};
    const DWORD written          = GetModuleFileNameW(nullptr, modulePath, static_cast<DWORD>(std::size(modulePath)));
    state.Require(written > 0u && written < std::size(modulePath), L"ViewerPE close roundtrip: failed to resolve the current module path.");
    if (written == 0u || written >= std::size(modulePath))
    {
        return false;
    }

    return RunStandaloneViewerCloseRoundTrip(
        state, L"ViewerPE close roundtrip", L"builtin/viewer-pe", L"RedSalamander.ViewerPE", std::filesystem::path(modulePath));
}

[[nodiscard]] bool TestViewerImgRawCloseRoundTrip(CaseState& state) noexcept
{
    std::error_code ec;
    const std::filesystem::path tempRoot = std::filesystem::temp_directory_path(ec) / std::format(L"RedSalamander.ViewerImgRawClose.{}", NewGuidText());
    state.Require(! ec, L"ViewerImgRaw close roundtrip: failed to resolve the temp directory.");
    if (ec)
    {
        return false;
    }

    std::filesystem::create_directories(tempRoot, ec);
    state.Require(! ec, L"ViewerImgRaw close roundtrip: failed to create the temp directory.");
    if (ec)
    {
        return false;
    }

    auto cleanupTempRoot = wil::scope_exit([&]() noexcept
    {
        std::error_code cleanupEc;
        std::filesystem::remove_all(tempRoot, cleanupEc);
    });

    static constexpr auto kFixtureBytes = std::to_array<std::byte>({
        std::byte{0x52}, std::byte{0x53}, std::byte{0x01}, std::byte{0x02}, std::byte{0x10}, std::byte{0x20}, std::byte{0x30}, std::byte{0x40},
        std::byte{0x50}, std::byte{0x60}, std::byte{0x70}, std::byte{0x80}, std::byte{0x90}, std::byte{0xA0}, std::byte{0xB0}, std::byte{0xC0},
        std::byte{0xD0}, std::byte{0xE0}, std::byte{0xF0}, std::byte{0x0F}, std::byte{0x1E}, std::byte{0x2D}, std::byte{0x3C}, std::byte{0x4B},
        std::byte{0x5A}, std::byte{0x69}, std::byte{0x78}, std::byte{0x87}, std::byte{0x96}, std::byte{0xA5}, std::byte{0xB4}, std::byte{0xC3},
    });

    const std::filesystem::path samplePath = tempRoot / L"viewer-imgraw-close-selftest.raw";
    {
        wil::unique_handle file(CreateFileW(samplePath.c_str(), GENERIC_WRITE, 0u, nullptr, CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, nullptr));
        state.Require(file.is_valid(), L"ViewerImgRaw close roundtrip: failed to create the fixture.");
        if (! file)
        {
            return false;
        }

        DWORD written      = 0u;
        const BOOL writeOk = WriteFile(file.get(), kFixtureBytes.data(), static_cast<DWORD>(kFixtureBytes.size()), &written, nullptr);
        state.Require(writeOk != FALSE && written == kFixtureBytes.size(), L"ViewerImgRaw close roundtrip: failed to write the fixture.");
        if (writeOk == FALSE || written != kFixtureBytes.size())
        {
            return false;
        }
    }

    return RunStandaloneViewerCloseRoundTrip(state, L"ViewerImgRaw close roundtrip", L"builtin/viewer-imgraw", L"RedSalamander.ViewerImgRaw", samplePath);
}

[[nodiscard]] bool TestViewerWebCloseRoundTrip(CaseState& state) noexcept
{
    std::error_code ec;
    const std::filesystem::path tempRoot = std::filesystem::temp_directory_path(ec) / std::format(L"RedSalamander.ViewerWebClose.{}", NewGuidText());
    state.Require(! ec, L"ViewerWeb close roundtrip: failed to resolve the temp directory.");
    if (ec)
    {
        return false;
    }

    std::filesystem::create_directories(tempRoot, ec);
    state.Require(! ec, L"ViewerWeb close roundtrip: failed to create the temp directory.");
    if (ec)
    {
        return false;
    }

    auto cleanupTempRoot = wil::scope_exit([&]() noexcept
    {
        std::error_code cleanupEc;
        std::filesystem::remove_all(tempRoot, cleanupEc);
    });

    const std::filesystem::path samplePath = tempRoot / L"viewer-web-close-selftest.html";
    const std::string html                 = "<!doctype html><html><head><meta charset=\"utf-8\"><title>viewer-web-close</title></head>"
                                             "<body><h1>viewer-web-close</h1><p>DxUi teardown close selftest.</p></body></html>";
    state.Require(SelfTest::WriteTextFile(samplePath, html), L"ViewerWeb close roundtrip: failed to write the HTML fixture.");
    if (! state.failure.empty())
    {
        return false;
    }

    return RunStandaloneViewerCloseRoundTrip(state, L"ViewerWeb close roundtrip", L"builtin/viewer-web", L"RedSalamander.ViewerWeb", samplePath);
}

[[nodiscard]] std::wstring GetEnvVarTrimmed(std::wstring_view name) noexcept
{
    if (name.empty())
    {
        return {};
    }

    std::wstring key(name);
    const DWORD required = ::GetEnvironmentVariableW(key.c_str(), nullptr, 0u);
    if (required == 0u)
    {
        return {};
    }

    std::wstring value(required, L'\0');
    const DWORD written = ::GetEnvironmentVariableW(key.c_str(), value.data(), required);
    if (written == 0u || written >= required)
    {
        return {};
    }

    value.resize(written);
    const auto notSpace = [](wchar_t ch) noexcept { return ! std::iswspace(static_cast<wint_t>(ch)); };
    const auto first    = std::find_if(value.begin(), value.end(), notSpace);
    if (first == value.end())
    {
        return {};
    }

    const auto last = std::find_if(value.rbegin(), value.rend(), notSpace).base();
    return std::wstring(first, last);
}

[[nodiscard]] std::wstring MakeUniquePipeName() noexcept
{
    return std::format(LR"(\\.\pipe\RedSalamander.SearchService.Test.{})", NewGuidText());
}

[[nodiscard]] std::filesystem::path GetSiblingExecutablePath(std::wstring_view fileName) noexcept
{
    if (fileName.empty())
    {
        return {};
    }

    std::array<wchar_t, 32768> buffer{};
    const DWORD length = ::GetModuleFileNameW(nullptr, buffer.data(), static_cast<DWORD>(buffer.size()));
    if (length == 0u || length >= buffer.size())
    {
        return {};
    }

    std::filesystem::path path(buffer.data(), buffer.data() + length);
    return path.replace_filename(fileName);
}

class ForegroundSearchServiceProcess final
{
public:
    ForegroundSearchServiceProcess()                                                 = default;
    ForegroundSearchServiceProcess(const ForegroundSearchServiceProcess&)            = delete;
    ForegroundSearchServiceProcess& operator=(const ForegroundSearchServiceProcess&) = delete;
    ~ForegroundSearchServiceProcess()
    {
        Stop();
    }

    [[nodiscard]] bool Start(
        std::wstring_view pipeName, uint32_t maxRequestsBeforeExit, std::wstring_view extraArguments, bool waitForReady, std::wstring& outError) noexcept
    {
        outError.clear();

        const std::filesystem::path servicePath = GetSiblingExecutablePath(L"RedSalamanderSearchService.exe");
        std::error_code ec;
        if (servicePath.empty() || ! std::filesystem::exists(servicePath, ec))
        {
            outError = std::format(L"Service executable not found: {}", servicePath.wstring());
            return false;
        }

        _pipeName                = std::wstring(pipeName);
        std::wstring commandLine = std::format(L"\"{}\" --run-foreground --pipe-name=\"{}\" --max-requests={} --protocol-version={}",
                                               servicePath.wstring(),
                                               _pipeName,
                                               maxRequestsBeforeExit,
                                               SearchServiceBroker::kProtocolVersion);
        if (! extraArguments.empty())
        {
            commandLine.push_back(L' ');
            commandLine.append(extraArguments);
        }

        STARTUPINFOW startupInfo{};
        startupInfo.cb = sizeof(startupInfo);

        PROCESS_INFORMATION processInfo{};
        if (::CreateProcessW(nullptr, commandLine.data(), nullptr, nullptr, FALSE, 0u, nullptr, nullptr, &startupInfo, &processInfo) == 0)
        {
            outError = std::format(L"CreateProcessW failed. error={}", ::GetLastError());
            return false;
        }

        _process.reset(processInfo.hProcess);
        _thread.reset(processInfo.hThread);
        if (! waitForReady)
        {
            return true;
        }

        return WaitUntilReady(outError);
    }

    void Stop() noexcept
    {
        if (_process)
        {
            static_cast<void>(::TerminateProcess(_process.get(), 0u));
            static_cast<void>(::WaitForSingleObject(_process.get(), 2000u));
        }

        _thread.reset();
        _process.reset();
        _pipeName.clear();
    }

private:
    [[nodiscard]] bool WaitForPipeReady(std::wstring& outError, bool allowExitedProcess) noexcept
    {
        if (_pipeName.empty())
        {
            outError = L"Foreground search service pipe name is empty.";
            return false;
        }

        for (int attempt = 0; attempt < 50; ++attempt)
        {
            if (_process && ::WaitForSingleObject(_process.get(), 0u) == WAIT_OBJECT_0)
            {
                DWORD exitCode = 0u;
                static_cast<void>(::GetExitCodeProcess(_process.get(), &exitCode));
                if (allowExitedProcess && exitCode == 0u)
                {
                    return true;
                }

                outError = std::format(L"Service process exited before pipe readiness. exitCode={}", exitCode);
                return false;
            }

            if (::WaitNamedPipeW(_pipeName.c_str(), 50u) != 0)
            {
                return true;
            }

            const DWORD error = ::GetLastError();
            if (error != ERROR_FILE_NOT_FOUND && error != ERROR_SEM_TIMEOUT && error != ERROR_PIPE_BUSY)
            {
                outError = std::format(L"WaitNamedPipeW failed while waiting for the foreground search service pipe. error={}", error);
                return false;
            }

            std::this_thread::sleep_for(std::chrono::milliseconds(25));
        }

        outError = L"Timed out waiting for the foreground search service pipe.";
        return false;
    }

    [[nodiscard]] bool WaitUntilReady(std::wstring& outError) noexcept
    {
        if (! WaitForPipeReady(outError, false))
        {
            return false;
        }

        for (int attempt = 0; attempt < 50; ++attempt)
        {
            SearchServiceBroker::ServiceStatus status{};
            if (SUCCEEDED(SearchServiceBroker::GetStatus(status)))
            {
                return WaitForPipeReady(outError, true);
            }

            if (_process && ::WaitForSingleObject(_process.get(), 0u) == WAIT_OBJECT_0)
            {
                DWORD exitCode = 0u;
                static_cast<void>(::GetExitCodeProcess(_process.get(), &exitCode));
                outError = std::format(L"Service process exited before readiness. exitCode={}", exitCode);
                return false;
            }

            std::this_thread::sleep_for(std::chrono::milliseconds(50));
        }

        outError = L"Timed out waiting for the search service foreground process.";
        return false;
    }

    wil::unique_handle _process;
    wil::unique_handle _thread;
    std::wstring _pipeName;
};

[[nodiscard]] bool TestPluginConfigurationDialogUsesDxUiFormSurface(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    const FileSystemPluginManager::PluginEntry* entry = FindFileSystemPluginById(kBuiltinLocalFileSystemId);
    state.Require(entry != nullptr, L"builtin/file-system plugin entry unavailable for plugin configuration self-test.");
    if (! entry)
    {
        return false;
    }

    struct CycleResult
    {
        bool sawDialog                 = false;
        bool ownedByMainWindow         = false;
        bool capturedSnapshot          = false;
        bool cancelled                 = false;
        size_t visibleUiaProviderCount = 0u;
        std::optional<UiaDescendantPatternStats> patternStats;
        std::optional<UiaValuePatternState> valueState;
        std::optional<UiaTogglePatternState> toggleState;
        PluginConfigurationDialogDebugSnapshot snapshot{};
    };

    const auto baselineSnapshotReady = [](const PluginConfigurationDialogDebugSnapshot& value) noexcept
    {
        return value.usesDxUiCommandButtons && value.usesDxUiFormSurface && value.usesDxUiFormStatics && value.usesDxUiFormInputs &&
               value.legacyOwnerDrawCommandButtonCount == 0u && value.legacyOwnerDrawFormInputCount == 0u && value.visibleLegacyCommandButtonCount == 0u &&
               value.visibleLegacyFormStaticCount == 0u && value.visibleLegacyFormInputCount == 0u && value.visibleDxCommandButtonHostCount == 2u &&
               value.visibleDxFormStaticHostCount > 0u && value.visibleDxFormInputHostCount > 0u;
    };

    const auto validateCycle = [&](const CycleResult& cycleResult, std::wstring_view context) noexcept
    {
        state.Require(cycleResult.sawDialog, std::format(L"Plugin configuration dialog did not open during {}.", context));
        state.Require(cycleResult.ownedByMainWindow, std::format(L"Plugin configuration dialog should be owned by the main window during {}.", context));
        state.Require(cycleResult.capturedSnapshot, std::format(L"Failed to capture plugin configuration dialog snapshot during {}.", context));
        if (! state.failure.empty())
        {
            return false;
        }

        state.Require(cycleResult.snapshot.usesDxUiCommandButtons,
                      std::format(L"Plugin configuration dialog should use shared DxUi command buttons during {}.", context));
        state.Require(
            cycleResult.snapshot.legacyOwnerDrawCommandButtonCount == 0u,
            std::format(L"Plugin configuration dialog should not keep hidden legacy owner-draw command buttons behind the DxUi shell during {}.", context));
        state.Require(
            cycleResult.snapshot.legacyOwnerDrawFormInputCount == 0u,
            std::format(L"Plugin configuration dialog should not keep hidden legacy owner-draw toggle inputs behind the DxUi form during {}.", context));
        state.Require(cycleResult.snapshot.visibleLegacyCommandButtonCount == 0u,
                      std::format(L"Plugin configuration dialog should hide legacy OK/Cancel buttons during {}.", context));
        state.Require(cycleResult.snapshot.visibleDxCommandButtonHostCount == 2u,
                      std::format(L"Plugin configuration dialog should expose exactly two visible DxUi command-button hosts during {}.", context));
        state.Require(cycleResult.snapshot.usesDxUiFormSurface,
                      std::format(L"Plugin configuration dialog should expose the re-landed DxUi form surface during {}.", context));
        state.Require(cycleResult.snapshot.usesDxUiFormStatics, std::format(L"Plugin configuration dialog should use shared DxUi statics during {}.", context));
        state.Require(cycleResult.snapshot.usesDxUiFormInputs,
                      std::format(L"Plugin configuration dialog interactive inputs should use shared DxUi hosts during {}.", context));
        state.Require(cycleResult.snapshot.visibleLegacyFormStaticCount == 0u,
                      std::format(L"Plugin configuration dialog should hide visible legacy form statics during {}.", context));
        state.Require(cycleResult.snapshot.visibleLegacyFormInputCount == 0u,
                      std::format(L"Plugin configuration dialog should hide visible legacy form inputs during {}.", context));
        state.Require(cycleResult.snapshot.visibleDxFormStaticHostCount > 0u,
                      std::format(L"Plugin configuration dialog should expose visible DxUi form-static hosts during {}.", context));
        state.Require(cycleResult.snapshot.visibleDxFormInputHostCount > 0u,
                      std::format(L"Plugin configuration dialog should expose visible DxUi input hosts during {}.", context));

        const size_t expectedMinimumUiaProviderCount = cycleResult.snapshot.visibleDxCommandButtonHostCount +
                                                       cycleResult.snapshot.visibleDxFormStaticHostCount + cycleResult.snapshot.visibleDxFormInputHostCount;
        state.Require(cycleResult.visibleUiaProviderCount >= expectedMinimumUiaProviderCount,
                      std::format(L"Plugin configuration dialog should expose at least {} visible WM_GETOBJECT/UIA providers during {}; saw {}.",
                                  expectedMinimumUiaProviderCount,
                                  context,
                                  cycleResult.visibleUiaProviderCount));
        state.Require(cycleResult.patternStats.has_value(),
                      std::format(L"Failed to collect live UI Automation pattern statistics for the plugin configuration dialog during {}.", context));
        if (cycleResult.patternStats.has_value())
        {
            state.Require(cycleResult.patternStats->visibleElementCount > 0u,
                          std::format(L"Plugin configuration dialog should expose visible UI Automation descendants during {}.", context));
            state.Require(cycleResult.patternStats->editControlCount + cycleResult.patternStats->comboBoxControlCount +
                                  cycleResult.patternStats->checkBoxControlCount + cycleResult.patternStats->radioButtonControlCount >
                              0u,
                          std::format(L"Plugin configuration dialog should expose at least one visible UI Automation input descendant during {}.", context));
            state.Require(cycleResult.patternStats->valuePatternCount + cycleResult.patternStats->togglePatternCount +
                                  cycleResult.patternStats->rangeValuePatternCount >
                              0u,
                          std::format(L"Plugin configuration dialog should expose live UI Automation input patterns during {}.", context));
            state.Require(cycleResult.patternStats->valuePatternCount > 0u,
                          std::format(L"Plugin configuration dialog should expose live UI Automation ValuePattern support during {}.", context));
            state.Require(cycleResult.patternStats->togglePatternCount > 0u,
                          std::format(L"Plugin configuration dialog should expose live UI Automation TogglePattern support during {}.", context));
        }

        state.Require(cycleResult.valueState.has_value(),
                      std::format(L"Failed to collect plugin configuration visible DX edit ValuePattern state during {}.", context));
        if (cycleResult.valueState.has_value())
        {
            state.Require(! cycleResult.valueState->isReadOnly,
                          std::format(L"Plugin configuration visible DX edit surface should remain editable during {}.", context));
            state.Require(! cycleResult.valueState->name.empty(),
                          std::format(L"Plugin configuration visible DX edit surface should expose a stable accessible name during {}.", context));
        }

        state.Require(cycleResult.toggleState.has_value(), std::format(L"Failed to collect plugin configuration visible DX toggle state during {}.", context));
        if (cycleResult.toggleState.has_value())
        {
            state.Require(! cycleResult.toggleState->name.empty(),
                          std::format(L"Plugin configuration visible DX toggle should expose a stable accessible name during {}.", context));
        }

        state.Require(cycleResult.cancelled, std::format(L"Plugin configuration dialog debug cancel command failed during {}.", context));
        return state.failure.empty();
    };

    const auto runCycle = [&](std::wstring_view context, std::wstring_view themeName) noexcept
    {
        CycleResult cycleResult{};
        const size_t baselineAttachedWindowHostCount = RedSalamander::DxUi::DebugGetAttachedWindowHostCount();
        const auto waitForAttachedWindowHostCount    = [&](const size_t expectedCount, const auto timeout) noexcept
        {
            const auto deadline = std::chrono::steady_clock::now() + timeout;
            while (std::chrono::steady_clock::now() < deadline)
            {
                PumpPendingMessages();
                if (RedSalamander::DxUi::DebugGetAttachedWindowHostCount() == expectedCount)
                {
                    return true;
                }
                std::this_thread::sleep_for(20ms);
            }

            return RedSalamander::DxUi::DebugGetAttachedWindowHostCount() == expectedCount;
        };
        std::jthread worker([&](std::stop_token) noexcept
        {
            const HWND dialog     = WaitForWindow([] noexcept { return GetPluginConfigurationDialogHandle(); }, SelfTest::Scale(5000ms));
            cycleResult.sawDialog = dialog != nullptr && IsWindow(dialog) != FALSE;
            if (! cycleResult.sawDialog)
            {
                return;
            }

            cycleResult.ownedByMainWindow = IsOwnedBy(dialog, mainWindow);

            const auto waitForSnapshot = [&](const auto& predicate, PluginConfigurationDialogDebugSnapshot& outSnapshot) noexcept
            {
                const auto deadline = std::chrono::steady_clock::now() + SelfTest::Scale(3000ms);
                while (std::chrono::steady_clock::now() < deadline)
                {
                    PumpPendingMessages();
                    outSnapshot = {};
                    if (DebugGetPluginConfigurationDialogSnapshot(outSnapshot) && predicate(outSnapshot))
                    {
                        return true;
                    }
                    std::this_thread::sleep_for(20ms);
                }

                outSnapshot = {};
                return DebugGetPluginConfigurationDialogSnapshot(outSnapshot) && predicate(outSnapshot);
            };

            cycleResult.capturedSnapshot = waitForSnapshot(baselineSnapshotReady, cycleResult.snapshot);
            if (! cycleResult.capturedSnapshot && baselineSnapshotReady(cycleResult.snapshot))
            {
                cycleResult.capturedSnapshot = true;
            }
            if (! cycleResult.capturedSnapshot)
            {
                SelfTest::AppendSuiteTrace(
                    SelfTest::SelfTestSuite::Commands,
                    std::format(L"plugin-config baseline snapshot miss: usesButtons={} usesSurface={} usesStatics={} usesInputs={} legacyButtons={} "
                                L"legacyStatics={} legacyInputs={} dxButtons={} dxStatics={} dxInputs={} panelHeight={} scrollY={}",
                                cycleResult.snapshot.usesDxUiCommandButtons ? 1 : 0,
                                cycleResult.snapshot.usesDxUiFormSurface ? 1 : 0,
                                cycleResult.snapshot.usesDxUiFormStatics ? 1 : 0,
                                cycleResult.snapshot.usesDxUiFormInputs ? 1 : 0,
                                cycleResult.snapshot.visibleLegacyCommandButtonCount,
                                cycleResult.snapshot.visibleLegacyFormStaticCount,
                                cycleResult.snapshot.visibleLegacyFormInputCount,
                                cycleResult.snapshot.visibleDxCommandButtonHostCount,
                                cycleResult.snapshot.visibleDxFormStaticHostCount,
                                cycleResult.snapshot.visibleDxFormInputHostCount,
                                cycleResult.snapshot.panelClientHeight,
                                cycleResult.snapshot.panelScrollPosY));
            }
            if (cycleResult.capturedSnapshot)
            {
                const size_t expectedMinimumUiaProviderCount = cycleResult.snapshot.visibleDxCommandButtonHostCount +
                                                               cycleResult.snapshot.visibleDxFormStaticHostCount +
                                                               cycleResult.snapshot.visibleDxFormInputHostCount;
                const auto deadline                          = std::chrono::steady_clock::now() + SelfTest::Scale(1500ms);
                while (std::chrono::steady_clock::now() < deadline)
                {
                    PumpPendingMessages();
                    cycleResult.visibleUiaProviderCount = CountVisibleDescendantWindowsExposingUiaProviders(dialog);
                    cycleResult.patternStats            = CollectVisibleUiaDescendantPatternStats(dialog);
                    cycleResult.valueState              = CollectVisibleDescendantValuePatternState(dialog, UIA_EditControlTypeId);
                    cycleResult.toggleState             = CollectVisibleDescendantTogglePatternState(dialog);
                    if (cycleResult.visibleUiaProviderCount >= expectedMinimumUiaProviderCount && cycleResult.patternStats.has_value() &&
                        cycleResult.valueState.has_value() && cycleResult.toggleState.has_value())
                    {
                        break;
                    }
                    std::this_thread::sleep_for(20ms);
                }
            }
            cycleResult.cancelled = DebugCancelPluginConfigurationDialog();
        });

        Common::Settings::Settings baselineSettings = g_settings;
        Common::Settings::Settings workingSettings  = baselineSettings;
        const AppTheme theme                        = ResolveAppTheme(ThemeMode::Dark, themeName);
        const HRESULT hr =
            EditPluginConfigurationDialog(mainWindow, PluginType::FileSystem, kBuiltinLocalFileSystemId, entry->name, baselineSettings, workingSettings, theme);
        worker.join();

        state.Require(hr == S_FALSE,
                      std::format(L"EditPluginConfigurationDialog returned unexpected HRESULT 0x{:08X} during {}.", static_cast<uint32_t>(hr), context));
        state.Require(waitForAttachedWindowHostCount(baselineAttachedWindowHostCount, SelfTest::Scale(3000ms)),
                      std::format(L"Plugin configuration dialog left {} attached DxUI hosts after {}; expected baseline {}.",
                                  RedSalamander::DxUi::DebugGetAttachedWindowHostCount(),
                                  context,
                                  baselineAttachedWindowHostCount));
        if (! state.failure.empty())
        {
            return false;
        }

        state.Require(validateCycle(cycleResult, context),
                      std::format(L"Plugin configuration dialog baseline DX surface validation failed during {}.", context));
        return state.failure.empty();
    };

    state.Require(runCycle(L"initial baseline surface probe", L"plugin-config-selftest-initial"),
                  L"Initial plugin configuration baseline DX surface probe failed.");
    if (! state.failure.empty())
    {
        return false;
    }

    state.Require(runCycle(L"reopened baseline surface probe", L"plugin-config-selftest-reopened"),
                  L"Reopened plugin configuration baseline DX surface probe failed.");
    return state.failure.empty();
}

[[nodiscard]] bool TestPluginConfigurationDialogLongRunScrollingKeepsDxSurfaceStable(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    const FileSystemPluginManager::PluginEntry* entry = FindFileSystemPluginById(kBuiltinLocalFileSystemId);
    state.Require(entry != nullptr, L"builtin/file-system plugin entry unavailable for plugin configuration scrolling self-test.");
    if (! entry)
    {
        return false;
    }

    struct WorkerResult
    {
        bool sawDialog                = false;
        bool ownedByMainWindow        = false;
        bool capturedBaselineSnapshot = false;
        bool capturedFinalSnapshot    = false;
        bool sawScrollableOverflow    = false;
        bool scrolledDown             = false;
        bool restoredToTop            = false;
        bool cancelled                = false;
        std::optional<UiaDescendantPatternStats> patternStats;
        PluginConfigurationDialogDebugSnapshot baselineSnapshot{};
        PluginConfigurationDialogDebugSnapshot finalSnapshot{};
    } workerResult{};

    const auto baselineSnapshotReady = [](const PluginConfigurationDialogDebugSnapshot& value) noexcept
    {
        return value.usesDxUiCommandButtons && value.usesDxUiFormSurface && value.usesDxUiFormStatics && value.usesDxUiFormInputs &&
               value.visibleLegacyCommandButtonCount == 0u && value.visibleLegacyFormControlCount == 0u && value.visibleDxCommandButtonHostCount == 2u &&
               value.visibleDxFormStaticHostCount > 0u && value.visibleDxFormInputHostCount > 0u && value.panelContentHeight > 0 && value.panelClientHeight > 0;
    };

    std::jthread worker([&](std::stop_token) noexcept
    {
        const HWND dialog      = WaitForWindow([] noexcept { return GetPluginConfigurationDialogHandle(); }, SelfTest::Scale(5000ms));
        workerResult.sawDialog = dialog != nullptr && IsWindow(dialog) != FALSE;
        if (! workerResult.sawDialog)
        {
            return;
        }

        workerResult.ownedByMainWindow = IsOwnedBy(dialog, mainWindow);

        const auto waitForSnapshot = [&](const auto& predicate, PluginConfigurationDialogDebugSnapshot& outSnapshot) noexcept
        {
            const auto deadline = std::chrono::steady_clock::now() + SelfTest::Scale(3000ms);
            while (std::chrono::steady_clock::now() < deadline)
            {
                outSnapshot = {};
                if (DebugGetPluginConfigurationDialogSnapshot(outSnapshot) && predicate(outSnapshot))
                {
                    return true;
                }
                std::this_thread::sleep_for(20ms);
            }

            outSnapshot = {};
            return DebugGetPluginConfigurationDialogSnapshot(outSnapshot) && predicate(outSnapshot);
        };

        PluginConfigurationDialogDebugSnapshot snapshot{};
        workerResult.capturedBaselineSnapshot = waitForSnapshot(baselineSnapshotReady, snapshot);
        if (! workerResult.capturedBaselineSnapshot && baselineSnapshotReady(snapshot))
        {
            workerResult.capturedBaselineSnapshot = true;
        }
        if (! workerResult.capturedBaselineSnapshot)
        {
            workerResult.cancelled = DebugCancelPluginConfigurationDialog();
            return;
        }

        workerResult.baselineSnapshot      = snapshot;
        const int maxScroll                = std::max(0, snapshot.panelContentHeight - snapshot.panelClientHeight);
        workerResult.sawScrollableOverflow = snapshot.panelHasVerticalScrollbar && maxScroll > 0;
        if (! workerResult.sawScrollableOverflow)
        {
            workerResult.finalSnapshot         = snapshot;
            workerResult.capturedFinalSnapshot = true;
            workerResult.cancelled             = DebugCancelPluginConfigurationDialog();
            return;
        }

        int previousScrollPos = snapshot.panelScrollPosY;
        for (size_t chunk = 0; chunk < 8u && previousScrollPos < maxScroll; ++chunk)
        {
            if (! DebugScrollPluginConfigurationDialogByWheelDetents(-2))
            {
                break;
            }

            if (! waitForSnapshot([&](const PluginConfigurationDialogDebugSnapshot& value) noexcept { return value.panelScrollPosY > previousScrollPos; },
                                  snapshot))
            {
                break;
            }

            previousScrollPos = snapshot.panelScrollPosY;
        }

        workerResult.scrolledDown = previousScrollPos > workerResult.baselineSnapshot.panelScrollPosY;
        if (workerResult.scrolledDown)
        {
            while (previousScrollPos > 0)
            {
                if (! DebugScrollPluginConfigurationDialogByWheelDetents(2))
                {
                    break;
                }

                if (! waitForSnapshot([&](const PluginConfigurationDialogDebugSnapshot& value) noexcept { return value.panelScrollPosY < previousScrollPos; },
                                      snapshot))
                {
                    break;
                }

                previousScrollPos = snapshot.panelScrollPosY;
            }
        }

        workerResult.restoredToTop         = previousScrollPos == 0;
        workerResult.finalSnapshot         = snapshot;
        workerResult.capturedFinalSnapshot = true;
        workerResult.patternStats          = CollectVisibleUiaDescendantPatternStats(dialog);
        workerResult.cancelled             = DebugCancelPluginConfigurationDialog();
    });

    Common::Settings::Settings baselineSettings = g_settings;
    Common::Settings::Settings workingSettings  = baselineSettings;
    const AppTheme theme                        = ResolveAppTheme(ThemeMode::Dark, L"plugin-config-scroll-selftest");
    const HRESULT hr =
        EditPluginConfigurationDialog(mainWindow, PluginType::FileSystem, kBuiltinLocalFileSystemId, entry->name, baselineSettings, workingSettings, theme);
    worker.join();

    state.Require(workerResult.sawDialog, L"Plugin configuration dialog did not open for long-run scrolling validation.");
    state.Require(workerResult.ownedByMainWindow, L"Plugin configuration dialog should be owned by the main window during long-run scrolling validation.");
    state.Require(workerResult.capturedBaselineSnapshot, L"Failed to capture plugin configuration dialog baseline snapshot.");
    state.Require(workerResult.sawScrollableOverflow,
                  std::format(L"Plugin configuration dialog should expose vertical overflow for long-run scrolling validation; contentHeight={}, "
                              L"clientHeight={}, hasScrollbar={}.",
                              workerResult.baselineSnapshot.panelContentHeight,
                              workerResult.baselineSnapshot.panelClientHeight,
                              workerResult.baselineSnapshot.panelHasVerticalScrollbar ? 1 : 0));
    state.Require(workerResult.scrolledDown, L"Plugin configuration dialog did not scroll down on the active DxUi form surface.");
    state.Require(workerResult.restoredToTop, L"Plugin configuration dialog did not restore its scroll panel to the top after long-run scrolling.");
    state.Require(workerResult.capturedFinalSnapshot, L"Failed to capture plugin configuration dialog final snapshot after long-run scrolling.");
    state.Require(workerResult.finalSnapshot.usesDxUiCommandButtons && workerResult.finalSnapshot.usesDxUiFormSurface &&
                      workerResult.finalSnapshot.usesDxUiFormStatics && workerResult.finalSnapshot.usesDxUiFormInputs,
                  L"Plugin configuration dialog lost the active DxUi form surface during long-run scrolling.");
    state.Require(workerResult.finalSnapshot.visibleLegacyCommandButtonCount == 0u && workerResult.finalSnapshot.visibleLegacyFormControlCount == 0u,
                  L"Plugin configuration dialog exposed legacy visible chrome during long-run scrolling.");
    state.Require(workerResult.finalSnapshot.visibleDxCommandButtonHostCount == workerResult.baselineSnapshot.visibleDxCommandButtonHostCount,
                  std::format(L"Plugin configuration dialog command-button host count changed during scrolling; baseline={} final={}.",
                              workerResult.baselineSnapshot.visibleDxCommandButtonHostCount,
                              workerResult.finalSnapshot.visibleDxCommandButtonHostCount));
    state.Require(workerResult.finalSnapshot.visibleDxFormStaticHostCount == workerResult.baselineSnapshot.visibleDxFormStaticHostCount,
                  std::format(L"Plugin configuration dialog DxUi static-host count changed during scrolling; baseline={} final={}.",
                              workerResult.baselineSnapshot.visibleDxFormStaticHostCount,
                              workerResult.finalSnapshot.visibleDxFormStaticHostCount));
    state.Require(workerResult.finalSnapshot.visibleDxFormInputHostCount == workerResult.baselineSnapshot.visibleDxFormInputHostCount,
                  std::format(L"Plugin configuration dialog DxUi input-host count changed during scrolling; baseline={} final={}.",
                              workerResult.baselineSnapshot.visibleDxFormInputHostCount,
                              workerResult.finalSnapshot.visibleDxFormInputHostCount));
    state.Require(workerResult.finalSnapshot.panelHasVerticalScrollbar,
                  L"Plugin configuration dialog lost its vertical scrollbar during long-run scrolling validation.");
    state.Require(workerResult.finalSnapshot.panelScrollPosY == 0,
                  std::format(L"Plugin configuration dialog should finish long-run scrolling validation at the top; saw scrollPosY={}.",
                              workerResult.finalSnapshot.panelScrollPosY));
    state.Require(workerResult.patternStats.has_value(), L"Failed to collect plugin configuration dialog UI Automation patterns after long-run scrolling.");
    if (workerResult.patternStats.has_value())
    {
        state.Require(workerResult.patternStats->visibleElementCount > 0u,
                      L"Plugin configuration dialog should keep visible UI Automation descendants after long-run scrolling.");
        state.Require(workerResult.patternStats->valuePatternCount + workerResult.patternStats->togglePatternCount +
                              workerResult.patternStats->rangeValuePatternCount >
                          0u,
                      L"Plugin configuration dialog should keep live UI Automation input patterns after long-run scrolling.");
    }
    state.Require(workerResult.cancelled, L"Plugin configuration dialog debug cancel command failed after long-run scrolling validation.");
    state.Require(
        hr == S_FALSE,
        std::format(L"EditPluginConfigurationDialog returned unexpected HRESULT 0x{:08X} during long-run scrolling validation.", static_cast<uint32_t>(hr)));
    return state.failure.empty();
}

[[nodiscard]] bool TestPluginConfigurationDialogLongRunOpenCloseStaysStable(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    const FileSystemPluginManager::PluginEntry* entry = FindFileSystemPluginById(kBuiltinLocalFileSystemId);
    state.Require(entry != nullptr, L"builtin/file-system plugin entry unavailable for plugin configuration churn self-test.");
    if (! entry)
    {
        return false;
    }

    const auto closeExistingDialog = [&]() noexcept
    {
        if (const HWND existing = GetPluginConfigurationDialogHandle(); existing && IsWindow(existing) != FALSE)
        {
            PostMessageW(existing, WM_CLOSE, 0, 0);
            static_cast<void>(WaitForWindowClosed(existing, SelfTest::Scale(3000ms)));
        }
    };

    closeExistingDialog();

    constexpr size_t kCycles         = 12u;
    const auto baselineSnapshotReady = [](const PluginConfigurationDialogDebugSnapshot& value) noexcept
    {
        return value.usesDxUiCommandButtons && value.usesDxUiFormSurface && value.usesDxUiFormStatics && value.usesDxUiFormInputs &&
               value.visibleLegacyCommandButtonCount == 0u && value.visibleLegacyFormControlCount == 0u && value.visibleDxCommandButtonHostCount == 2u &&
               value.visibleDxFormInputHostCount > 0u;
    };
    for (size_t cycle = 0; cycle < kCycles; ++cycle)
    {
        const bool accept = (cycle % 2u) == 0u;

        struct WorkerResult
        {
            HWND dialog                    = nullptr;
            bool sawDialog                 = false;
            bool ownedByMainWindow         = false;
            bool capturedSnapshot          = false;
            bool closeIssued               = false;
            size_t visibleUiaProviderCount = 0u;
            std::optional<UiaDescendantPatternStats> patternStats;
            std::optional<UiaValuePatternState> valueState;
            std::optional<UiaTogglePatternState> toggleState;
            std::wstring buttonName;
            PluginConfigurationDialogDebugSnapshot snapshot{};
        } workerResult{};

        std::jthread worker([&](std::stop_token) noexcept
        {
            workerResult.dialog    = WaitForWindow([] noexcept { return GetPluginConfigurationDialogHandle(); }, SelfTest::Scale(5000ms));
            workerResult.sawDialog = workerResult.dialog != nullptr && IsWindow(workerResult.dialog) != FALSE;
            if (! workerResult.sawDialog)
            {
                return;
            }

            workerResult.ownedByMainWindow = IsOwnedBy(workerResult.dialog, mainWindow);
            workerResult.capturedSnapshot  = DebugGetPluginConfigurationDialogSnapshot(workerResult.snapshot);
            if (! workerResult.capturedSnapshot || ! baselineSnapshotReady(workerResult.snapshot))
            {
                const auto deadline = std::chrono::steady_clock::now() + SelfTest::Scale(3000ms);
                while (std::chrono::steady_clock::now() < deadline)
                {
                    PumpPendingMessages();
                    workerResult.snapshot = {};
                    if (DebugGetPluginConfigurationDialogSnapshot(workerResult.snapshot) && baselineSnapshotReady(workerResult.snapshot))
                    {
                        workerResult.capturedSnapshot = true;
                        break;
                    }
                    std::this_thread::sleep_for(20ms);
                }

                if (! workerResult.capturedSnapshot && baselineSnapshotReady(workerResult.snapshot))
                {
                    workerResult.capturedSnapshot = true;
                }
            }
            workerResult.visibleUiaProviderCount = CountVisibleDescendantWindowsExposingUiaProviders(workerResult.dialog);
            workerResult.patternStats            = CollectVisibleUiaDescendantPatternStats(workerResult.dialog);
            workerResult.valueState              = CollectVisibleDescendantValuePatternState(workerResult.dialog, UIA_EditControlTypeId);
            workerResult.toggleState             = CollectVisibleDescendantTogglePatternState(workerResult.dialog);
            if (const auto buttonState = CollectVisibleDescendantNamedElementState(workerResult.dialog, UIA_ButtonControlTypeId); buttonState.has_value())
            {
                workerResult.buttonName = buttonState->name;
            }

            if (accept)
            {
                workerResult.closeIssued = PostMessageW(workerResult.dialog, WM_COMMAND, MAKEWPARAM(IDOK, 0), 0) != FALSE;
            }
            else
            {
                workerResult.closeIssued = DebugCancelPluginConfigurationDialog();
            }
        });

        Common::Settings::Settings baselineSettings = g_settings;
        Common::Settings::Settings workingSettings  = baselineSettings;
        const AppTheme theme                        = ResolveAppTheme(ThemeMode::Dark, std::format(L"plugin-config-churn-{}", cycle));
        const HRESULT hr =
            EditPluginConfigurationDialog(mainWindow, PluginType::FileSystem, kBuiltinLocalFileSystemId, entry->name, baselineSettings, workingSettings, theme);
        worker.join();

        state.Require(workerResult.sawDialog, std::format(L"Plugin configuration dialog did not open during cycle {}.", cycle));
        state.Require(workerResult.ownedByMainWindow, std::format(L"Plugin configuration dialog should be owned by the main window during cycle {}.", cycle));
        state.Require(workerResult.capturedSnapshot, std::format(L"Failed to capture plugin configuration dialog snapshot during cycle {}.", cycle));
        state.Require(workerResult.closeIssued, std::format(L"Plugin configuration dialog did not accept the close action during cycle {}.", cycle));
        state.Require(workerResult.snapshot.usesDxUiCommandButtons,
                      std::format(L"Plugin configuration dialog should use DxUi command buttons during cycle {}.", cycle));
        state.Require(workerResult.snapshot.usesDxUiFormSurface,
                      std::format(L"Plugin configuration dialog should keep the DxUi form surface active during cycle {}.", cycle));
        state.Require(workerResult.snapshot.usesDxUiFormStatics,
                      std::format(L"Plugin configuration dialog should keep DxUi schema statics active during cycle {}.", cycle));
        state.Require(workerResult.snapshot.usesDxUiFormInputs,
                      std::format(L"Plugin configuration dialog should keep DxUi schema inputs active during cycle {}.", cycle));
        state.Require(workerResult.snapshot.visibleLegacyCommandButtonCount == 0u,
                      std::format(L"Plugin configuration dialog should not expose visible legacy command buttons during cycle {}; saw {}.",
                                  cycle,
                                  workerResult.snapshot.visibleLegacyCommandButtonCount));
        state.Require(workerResult.snapshot.visibleLegacyFormControlCount == 0u,
                      std::format(L"Plugin configuration dialog should not expose visible legacy form controls during cycle {}; saw {}.",
                                  cycle,
                                  workerResult.snapshot.visibleLegacyFormControlCount));
        state.Require(workerResult.snapshot.visibleDxCommandButtonHostCount == 2u,
                      std::format(L"Plugin configuration dialog should expose exactly two visible DxUi command-button hosts during cycle {}; saw {}.",
                                  cycle,
                                  workerResult.snapshot.visibleDxCommandButtonHostCount));
        state.Require(workerResult.snapshot.visibleDxFormStaticHostCount > 0u,
                      std::format(L"Plugin configuration dialog should expose visible DxUi schema statics during cycle {}.", cycle));
        state.Require(workerResult.snapshot.visibleDxFormInputHostCount > 0u,
                      std::format(L"Plugin configuration dialog should expose visible DxUi schema inputs during cycle {}.", cycle));

        const size_t expectedMinimumUiaProviderCount = workerResult.snapshot.visibleDxCommandButtonHostCount +
                                                       workerResult.snapshot.visibleDxFormStaticHostCount + workerResult.snapshot.visibleDxFormInputHostCount;
        state.Require(workerResult.visibleUiaProviderCount >= expectedMinimumUiaProviderCount,
                      std::format(L"Plugin configuration dialog should expose at least {} visible WM_GETOBJECT/UIA providers during cycle {}; saw {}.",
                                  expectedMinimumUiaProviderCount,
                                  cycle,
                                  workerResult.visibleUiaProviderCount));
        state.Require(workerResult.patternStats.has_value(),
                      std::format(L"Failed to collect plugin configuration dialog UI Automation patterns during cycle {}.", cycle));
        if (workerResult.patternStats.has_value())
        {
            state.Require(workerResult.patternStats->visibleElementCount > 0u,
                          std::format(L"Plugin configuration dialog should expose visible UI Automation descendants during cycle {}.", cycle));
            state.Require(workerResult.patternStats->editControlCount + workerResult.patternStats->comboBoxControlCount +
                                  workerResult.patternStats->checkBoxControlCount + workerResult.patternStats->radioButtonControlCount >
                              0u,
                          std::format(L"Plugin configuration dialog should expose a visible UI Automation input descendant during cycle {}.", cycle));
            state.Require(workerResult.patternStats->buttonControlCount > 0u,
                          std::format(L"Plugin configuration dialog should expose visible UI Automation command buttons during cycle {}.", cycle));
            state.Require(workerResult.patternStats->valuePatternCount + workerResult.patternStats->togglePatternCount +
                                  workerResult.patternStats->rangeValuePatternCount >
                              0u,
                          std::format(L"Plugin configuration dialog should expose live UI Automation input patterns during cycle {}.", cycle));
        }

        state.Require(workerResult.valueState.has_value(),
                      std::format(L"Failed to collect plugin configuration visible DX edit ValuePattern state during cycle {}.", cycle));
        if (workerResult.valueState.has_value())
        {
            state.Require(! workerResult.valueState->isReadOnly,
                          std::format(L"Plugin configuration visible DX edit surface should remain editable during cycle {}.", cycle));
            state.Require(! workerResult.valueState->name.empty(),
                          std::format(L"Plugin configuration visible DX edit surface should expose a stable accessible name during cycle {}.", cycle));
        }

        state.Require(workerResult.toggleState.has_value(),
                      std::format(L"Failed to collect plugin configuration visible DX toggle state during cycle {}.", cycle));
        if (workerResult.toggleState.has_value())
        {
            state.Require(! workerResult.toggleState->name.empty(),
                          std::format(L"Plugin configuration visible DX toggle should expose a stable accessible name during cycle {}.", cycle));
        }

        state.Require(! workerResult.buttonName.empty(),
                      std::format(L"Plugin configuration visible DX command button should expose a stable accessible name during cycle {}.", cycle));

        state.Require(hr == (accept ? S_OK : S_FALSE),
                      std::format(L"EditPluginConfigurationDialog returned unexpected HRESULT 0x{:08X} during {} cycle {}.",
                                  static_cast<uint32_t>(hr),
                                  accept ? L"accept" : L"cancel",
                                  cycle));

        const HWND lingeringDialog = GetPluginConfigurationDialogHandle();
        state.Require(lingeringDialog == nullptr || IsWindow(lingeringDialog) == FALSE,
                      std::format(L"Plugin configuration dialog should not remain open after cycle {}.", cycle));
        if (! state.failure.empty())
        {
            closeExistingDialog();
            return false;
        }
    }

    closeExistingDialog();
    return state.failure.empty();
}

[[nodiscard]] bool TestPluginConfigurationDialogThemeCycleKeepsSurfaceLegible(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;
    constexpr std::wstring_view kBuiltinS3FileSystemId = L"builtin/file-system-s3";
    SelfTest::AppendSelfTestTrace(L"PluginConfig theme-cycle: begin");
    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    const FileSystemPluginManager::PluginEntry* entry = FindFileSystemPluginById(kBuiltinS3FileSystemId);
    state.Require(entry != nullptr, L"builtin/file-system-s3 plugin entry unavailable for plugin configuration theme-cycle validation.");
    if (! entry)
    {
        return false;
    }

    struct WorkerResult
    {
        bool sawDialog                = false;
        bool ownedByMainWindow        = false;
        bool capturedBaselineSnapshot = false;
        bool cancelled                = false;
        bool closedAfterCancel        = false;
        PluginConfigurationDialogDebugSnapshot baselineSnapshot{};
    } workerResult{};

    std::jthread worker([&](std::stop_token) noexcept
    {
        SelfTest::AppendSelfTestTrace(L"PluginConfig theme-cycle: worker waiting for dialog");
        const HWND dialog      = WaitForWindow([] noexcept { return GetPluginConfigurationDialogHandle(); }, SelfTest::Scale(5000ms));
        workerResult.sawDialog = dialog != nullptr && IsWindow(dialog) != FALSE;
        if (! workerResult.sawDialog)
        {
            return;
        }
        SelfTest::AppendSelfTestTrace(std::format(L"PluginConfig theme-cycle: dialog opened hwnd=0x{:X}", reinterpret_cast<UINT_PTR>(dialog)));

        workerResult.ownedByMainWindow = IsOwnedBy(dialog, mainWindow);

        const auto waitForSnapshot = [&](const auto& predicate, PluginConfigurationDialogDebugSnapshot& outSnapshot) noexcept
        {
            const auto deadline = std::chrono::steady_clock::now() + SelfTest::Scale(3000ms);
            while (std::chrono::steady_clock::now() < deadline)
            {
                PumpPendingMessages();
                outSnapshot = {};
                if (DebugGetPluginConfigurationDialogSnapshot(outSnapshot) && predicate(outSnapshot))
                {
                    return true;
                }
                std::this_thread::sleep_for(20ms);
            }

            outSnapshot = {};
            return DebugGetPluginConfigurationDialogSnapshot(outSnapshot) && predicate(outSnapshot);
        };

        workerResult.capturedBaselineSnapshot = waitForSnapshot(
            [](const PluginConfigurationDialogDebugSnapshot& value) noexcept
        {
            return value.usesDxUiCommandButtons && value.usesDxUiFormSurface && value.usesDxUiFormStatics && value.usesDxUiFormInputs &&
                   value.visibleLegacyCommandButtonCount == 0u && value.visibleLegacyFormStaticCount == 0u && value.visibleLegacyFormInputCount == 0u &&
                   value.visibleDxCommandButtonHostCount == 2u && value.visibleDxFormStaticHostCount > 0u && value.visibleDxFormInputHostCount > 0u &&
                   value.themeDark && ! value.themeHighContrast && ! value.themeRainbow;
        },
            workerResult.baselineSnapshot);
        if (! workerResult.capturedBaselineSnapshot)
        {
            SelfTest::AppendSelfTestTrace(L"PluginConfig theme-cycle: baseline snapshot failed");
            static_cast<void>(DebugCancelPluginConfigurationDialog());
            return;
        }
        SelfTest::AppendSelfTestTrace(L"PluginConfig theme-cycle: baseline snapshot ready");

        state.Require(DebugFocusPluginConfigurationDialogFirstInput(),
                      L"Plugin configuration dialog did not focus the first visible DX input before theme-cycle validation.");
        PluginConfigurationDialogDebugSnapshot snapshot{};
        state.Require(waitForSnapshot(
                          [](const PluginConfigurationDialogDebugSnapshot& value) noexcept
        {
            return value.focusKind == PluginConfigurationDialogDebugFocusKind::Edit && value.focusLabel == L"Default region" &&
                   value.visibleDxFormInputHostCount > 0u;
        },
                          snapshot),
                      L"Plugin configuration dialog focus did not settle to the Default region edit before theme-cycle validation.");
        if (! state.failure.empty())
        {
            SelfTest::AppendSelfTestTrace(L"PluginConfig theme-cycle: initial focus failed");
            static_cast<void>(DebugCancelPluginConfigurationDialog());
            return;
        }
        SelfTest::AppendSelfTestTrace(L"PluginConfig theme-cycle: initial focus ready");

        const auto collectFocusedEditState = [&]() noexcept -> std::optional<UiaValuePatternState>
        {
            HWND focusedHost = nullptr;
            if (! DebugGetPluginConfigurationDialogFocusedHost(focusedHost) || ! focusedHost || IsWindow(focusedHost) == FALSE)
            {
                return std::nullopt;
            }

            return CollectWindowRootOrDescendantValuePatternState(focusedHost, UIA_EditControlTypeId);
        };

        const auto collectFirstVisibleToggleState = [&]() noexcept -> std::optional<UiaTogglePatternState>
        {
            HWND toggleHost = nullptr;
            RECT toggleRect{};
            if (! DebugGetPluginConfigurationDialogFirstVisibleToggleHostAndClientRect(toggleHost, toggleRect) || ! toggleHost ||
                IsWindow(toggleHost) == FALSE || toggleRect.right <= toggleRect.left || toggleRect.bottom <= toggleRect.top)
            {
                return std::nullopt;
            }

            return CollectWindowRootOrDescendantTogglePatternState(toggleHost);
        };

        const auto baselineEditState = collectFocusedEditState();
        state.Require(baselineEditState.has_value(), L"Plugin configuration dialog should expose a focused visible DX edit before theme-cycle validation.");
        if (! baselineEditState.has_value())
        {
            SelfTest::AppendSelfTestTrace(L"PluginConfig theme-cycle: baseline edit state missing");
            static_cast<void>(DebugCancelPluginConfigurationDialog());
            return;
        }

        state.Require(! baselineEditState->isReadOnly,
                      L"Plugin configuration dialog focused visible DX edit should remain editable during theme-cycle validation.");
        if (baselineEditState->isReadOnly)
        {
            SelfTest::AppendSelfTestTrace(L"PluginConfig theme-cycle: baseline edit readonly");
            static_cast<void>(DebugCancelPluginConfigurationDialog());
            return;
        }

        const auto baselineToggleState = collectFirstVisibleToggleState();
        state.Require(baselineToggleState.has_value(), L"Plugin configuration dialog should expose a visible DX toggle before theme-cycle validation.");
        if (! baselineToggleState.has_value())
        {
            SelfTest::AppendSelfTestTrace(L"PluginConfig theme-cycle: baseline toggle state missing");
            static_cast<void>(DebugCancelPluginConfigurationDialog());
            return;
        }

        const std::wstring baselineEditValue            = baselineEditState->value;
        const std::wstring baselineEditAccessibleName   = baselineEditState->name;
        const std::wstring baselineToggleAccessibleName = baselineToggleState->name;
        const ToggleState baselineToggleValue           = baselineToggleState->toggleState;

        const auto requireTheme = [&](std::wstring_view label, const AppTheme& theme, const bool expectRainbow, const bool expectHighContrast) noexcept
        {
            SelfTest::AppendSelfTestTrace(std::format(L"PluginConfig theme-cycle: apply {} theme begin", label));
            UpdatePluginConfigurationWindowsTheme(theme);
            state.Require(waitForSnapshot(
                              [&](const PluginConfigurationDialogDebugSnapshot& value) noexcept
            {
                return value.usesDxUiCommandButtons && value.usesDxUiFormSurface && value.usesDxUiFormStatics && value.usesDxUiFormInputs &&
                       value.visibleLegacyCommandButtonCount == 0u && value.visibleLegacyFormStaticCount == 0u && value.visibleLegacyFormInputCount == 0u &&
                       value.visibleDxCommandButtonHostCount == 2u && value.visibleDxFormStaticHostCount > 0u && value.visibleDxFormInputHostCount > 0u &&
                       value.themeDark == theme.dark && value.themeHighContrast == theme.highContrast && value.themeRainbow == theme.menu.rainbowMode;
            },
                              snapshot),
                          std::format(L"Plugin configuration dialog did not settle after the {} theme update.", label));
            if (! state.failure.empty())
            {
                SelfTest::AppendSelfTestTrace(std::format(L"PluginConfig theme-cycle: apply {} theme settle failed", label));
                return;
            }

            state.Require(DebugFocusPluginConfigurationDialogFirstInput(),
                          std::format(L"Plugin configuration dialog did not refocus the first visible DX input after the {} theme update.", label));
            state.Require(waitForSnapshot(
                              [](const PluginConfigurationDialogDebugSnapshot& value) noexcept
            {
                return value.focusKind == PluginConfigurationDialogDebugFocusKind::Edit && value.focusLabel == L"Default region" &&
                       value.visibleDxFormInputHostCount > 0u;
            },
                              snapshot),
                          std::format(L"Plugin configuration dialog focus did not return to the Default region edit after the {} theme update.", label));
            if (! state.failure.empty())
            {
                SelfTest::AppendSelfTestTrace(std::format(L"PluginConfig theme-cycle: apply {} theme refocus failed", label));
                return;
            }

            const auto valueState = collectFocusedEditState();
            state.Require(valueState.has_value(),
                          std::format(L"Plugin configuration dialog focused visible DX edit disappeared after the {} theme update.", label));
            if (valueState.has_value())
            {
                state.Require(! valueState->isReadOnly,
                              std::format(L"Plugin configuration dialog focused visible DX edit became read-only after the {} theme update.", label));
                state.Require(
                    valueState->name == baselineEditAccessibleName,
                    std::format(L"Plugin configuration dialog focused visible DX edit accessible name changed unexpectedly after the {} theme update.", label));
                state.Require(valueState->value == baselineEditValue,
                              std::format(L"Plugin configuration dialog focused visible DX edit value changed unexpectedly after the {} theme update.", label));
            }

            const auto toggleState = collectFirstVisibleToggleState();
            state.Require(toggleState.has_value(), std::format(L"Plugin configuration dialog visible DX toggle disappeared after the {} theme update.", label));
            if (toggleState.has_value())
            {
                state.Require(
                    toggleState->name == baselineToggleAccessibleName,
                    std::format(L"Plugin configuration dialog visible DX toggle accessible name changed unexpectedly after the {} theme update.", label));
                state.Require(toggleState->toggleState == baselineToggleValue,
                              std::format(L"Plugin configuration dialog visible DX toggle state changed unexpectedly after the {} theme update.", label));
            }

            state.Require(snapshot.themeRainbow == expectRainbow,
                          std::format(L"Plugin configuration dialog rainbow-theme flag mismatch after the {} theme update.", label));
            state.Require(snapshot.themeHighContrast == expectHighContrast,
                          std::format(L"Plugin configuration dialog high-contrast flag mismatch after the {} theme update.", label));
            if (state.failure.empty())
            {
                SelfTest::AppendSelfTestTrace(std::format(L"PluginConfig theme-cycle: apply {} theme complete", label));
            }
        };

        requireTheme(L"dark", ResolveAppTheme(ThemeMode::Dark, L"plugin-config-theme-cycle-dark"), false, false);
        requireTheme(L"light", ResolveAppTheme(ThemeMode::Light, L"plugin-config-theme-cycle-light"), false, false);
        requireTheme(L"rainbow", ResolveAppTheme(ThemeMode::Rainbow, L"plugin-config-theme-cycle-rainbow"), true, false);
        requireTheme(L"high-contrast", ResolveAppTheme(ThemeMode::HighContrast, L"plugin-config-theme-cycle-high-contrast"), false, true);

        workerResult.cancelled         = DebugCancelPluginConfigurationDialog();
        workerResult.closedAfterCancel = workerResult.cancelled && WaitForWindowClosed(dialog, SelfTest::Scale(3000ms));
        SelfTest::AppendSelfTestTrace(std::format(L"PluginConfig theme-cycle: cancel={} closed={}", workerResult.cancelled, workerResult.closedAfterCancel));
    });

    Common::Settings::Settings baselineSettings = g_settings;
    Common::Settings::Settings workingSettings  = baselineSettings;
    const AppTheme initialTheme                 = ResolveAppTheme(ThemeMode::Dark, L"plugin-config-theme-cycle-initial");
    const HRESULT hr =
        EditPluginConfigurationDialog(mainWindow, PluginType::FileSystem, kBuiltinS3FileSystemId, entry->name, baselineSettings, workingSettings, initialTheme);
    worker.join();

    state.Require(workerResult.sawDialog, L"Plugin configuration dialog did not open for theme-cycle validation.");
    state.Require(workerResult.ownedByMainWindow, L"Plugin configuration dialog should be owned by the main window during theme-cycle validation.");
    state.Require(workerResult.capturedBaselineSnapshot, L"Failed to capture the plugin configuration baseline snapshot before theme-cycle validation.");
    state.Require(workerResult.cancelled, L"Plugin configuration dialog did not accept the shared debug cancel path after theme-cycle validation.");
    state.Require(workerResult.closedAfterCancel, L"Plugin configuration dialog did not close cleanly after theme-cycle validation.");
    state.Require(
        hr == S_FALSE,
        std::format(L"EditPluginConfigurationDialog returned unexpected HRESULT 0x{:08X} during theme-cycle validation.", static_cast<unsigned int>(hr)));
    state.Require(GetPluginConfigurationDialogHandle() == nullptr || IsWindow(GetPluginConfigurationDialogHandle()) == FALSE,
                  L"Plugin configuration dialog should not remain open after theme-cycle validation.");
    return state.failure.empty();
}

[[nodiscard]] bool TestPluginConfigurationDialogLiveDxInteraction(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;
    constexpr std::wstring_view kBuiltinS3FileSystemId = L"builtin/file-system-s3";

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    const FileSystemPluginManager::PluginEntry* entry = FindFileSystemPluginById(kBuiltinS3FileSystemId);
    state.Require(entry != nullptr, L"builtin/file-system-s3 plugin entry unavailable for plugin configuration interaction self-test.");
    if (! entry)
    {
        return false;
    }

    const std::filesystem::path suiteRoot = SelfTest::GetTempRoot(SelfTest::SelfTestSuite::Commands);
    state.Require(! suiteRoot.empty(), L"SelfTest temp root unavailable for plugin configuration interaction test.");
    if (! state.failure.empty())
    {
        return false;
    }

    const std::filesystem::path browseTarget = suiteRoot / L"plugin_config_browse_target";
    std::error_code browseCreateError;
    std::filesystem::create_directories(browseTarget, browseCreateError);
    const std::string browseCreateErrorMessage = browseCreateError.message();
    state.Require(! browseCreateError,
                  std::format(L"Failed to create plugin configuration browse target directory: {}",
                              std::wstring(browseCreateErrorMessage.begin(), browseCreateErrorMessage.end())));
    state.Require(DebugSetPluginConfigurationNextBrowsePath(browseTarget.native()),
                  L"Failed to seed plugin configuration browse override for live interaction validation.");
    const auto resetBrowseOverride                   = wil::scope_exit([]() noexcept { static_cast<void>(DebugSetPluginConfigurationNextBrowsePath({})); });
    const std::filesystem::path liveStepArtifactPath = SelfTest::GetSuiteArtifactPath(SelfTest::SelfTestSuite::Commands, L"plugin_config_live_step.txt");
    static_cast<void>(SelfTest::WriteTextFile(liveStepArtifactPath, L"scheduled"));
    if (! state.failure.empty())
    {
        return false;
    }

    struct WorkerResult
    {
        bool sawDialog                        = false;
        bool ownedByMainWindow                = false;
        bool capturedBaselineSnapshot         = false;
        bool capturedFinalSnapshot            = false;
        bool sawReopenedDialog                = false;
        bool reopenedOwnedByMainWindow        = false;
        bool capturedReopenedSnapshot         = false;
        bool browseButtonAvailable            = false;
        bool sawBrowseButton                  = false;
        bool browseCancelPreservedVisibleEdit = false;
        bool browseUpdatedVisibleEdit         = false;
        bool browseRestoredAfterCancel        = false;
        bool mutatedEdit                      = false;
        bool restoredEdit                     = false;
        bool restoredToggle                   = false;
        bool toggleLabelChanged               = false;
        bool toggleLabelRestored              = false;
        bool reopenedEditRestored             = false;
        bool reopenedToggleRestored           = false;
        bool reopenedBrowseRoundTrip          = false;
        bool reopenedEditRoundTrip            = false;
        bool reopenedToggleRoundTrip          = false;
        bool reopenedToggleLabelRestored      = false;
        bool reopenedToggleLabelRoundTrip     = false;
        bool invokedCancel                    = false;
        bool closedAfterCancel                = false;
        bool invokedOk                        = false;
        bool closedAfterInvoke                = false;
        std::wstring blockedStep;
        std::wstring editDiagnostics;
        PluginConfigurationDialogDebugSnapshot baselineSnapshot{};
        PluginConfigurationDialogDebugSnapshot finalSnapshot{};
        PluginConfigurationDialogDebugSnapshot reopenedSnapshot{};
    } workerResult{};

    std::jthread worker([&](std::stop_token) noexcept
    {
        const auto traceStep = [&](std::wstring_view step) noexcept
        {
            workerResult.blockedStep.assign(step);
            SelfTest::AppendSuiteTrace(SelfTest::SelfTestSuite::Commands, std::format(L"plugin-config live: {}", step));
            static_cast<void>(SelfTest::WriteTextFile(liveStepArtifactPath, std::wstring(step)));
        };

        traceStep(L"wait initial dialog");
        const HWND dialog      = WaitForWindow([] noexcept { return GetPluginConfigurationDialogHandle(); }, SelfTest::Scale(5000ms));
        workerResult.sawDialog = dialog != nullptr && IsWindow(dialog) != FALSE;
        if (! workerResult.sawDialog)
        {
            return;
        }

        traceStep(L"initial dialog ready");
        workerResult.ownedByMainWindow = IsOwnedBy(dialog, mainWindow);

        const auto waitForSnapshot = [&](HWND expectedDialog, const auto& predicate, PluginConfigurationDialogDebugSnapshot& outSnapshot) noexcept
        {
            const auto deadline = std::chrono::steady_clock::now() + SelfTest::Scale(3000ms);
            while (std::chrono::steady_clock::now() < deadline)
            {
                PumpPendingMessages();
                if (GetPluginConfigurationDialogHandle() != expectedDialog)
                {
                    std::this_thread::sleep_for(20ms);
                    continue;
                }
                outSnapshot = {};
                if (DebugGetPluginConfigurationDialogSnapshot(outSnapshot) && predicate(outSnapshot))
                {
                    return true;
                }
                std::this_thread::sleep_for(20ms);
            }

            outSnapshot = {};
            return DebugGetPluginConfigurationDialogSnapshot(outSnapshot) && predicate(outSnapshot);
        };

        workerResult.capturedBaselineSnapshot = waitForSnapshot(dialog,
                                                                [](const PluginConfigurationDialogDebugSnapshot& value) noexcept
        {
            return value.usesDxUiCommandButtons && value.usesDxUiFormSurface && value.usesDxUiFormStatics && value.usesDxUiFormInputs &&
                   value.visibleLegacyCommandButtonCount == 0u && value.visibleLegacyFormControlCount == 0u && value.visibleDxCommandButtonHostCount == 2u &&
                   value.visibleDxFormInputHostCount > 0u;
        },
                                                                workerResult.baselineSnapshot);
        if (! workerResult.capturedBaselineSnapshot)
        {
            traceStep(L"baseline snapshot miss");
            static_cast<void>(DebugCancelPluginConfigurationDialog());
            return;
        }

        traceStep(L"baseline snapshot captured");

        std::wstring browseFieldName;
        std::wstring browseFieldBaselineValue;
        std::wstring editFieldName;
        std::wstring initialEditValue;
        std::wstring editedValue;
        std::wstring toggleName;
        std::optional<ToggleState> initialToggleValue;

        const auto collectVisibleEditableStates = [&](HWND targetDialog) noexcept
        {
            std::vector<UiaValuePatternState> states;
            for (const auto& element : FindMatchingVisibleDescendantElements(targetDialog, UIA_EditControlTypeId))
            {
                wil::com_ptr<IValueProvider> valueProvider;
                if (FAILED(element->GetCurrentPatternAs(UIA_ValuePatternId, __uuidof(IValueProvider), valueProvider.put_void())) || ! valueProvider)
                {
                    continue;
                }

                BOOL readOnly = FALSE;
                if (SUCCEEDED(valueProvider->get_IsReadOnly(&readOnly)) && readOnly != FALSE)
                {
                    continue;
                }

                UiaValuePatternState state{};
                state.controlType = UIA_EditControlTypeId;
                wil::unique_bstr name;
                if (SUCCEEDED(element->get_CurrentName(&name)))
                {
                    state.name.assign(name.get() ? name.get() : L"");
                }

                wil::unique_bstr value;
                if (SUCCEEDED(valueProvider->get_Value(&value)))
                {
                    state.value.assign(value.get() ? value.get() : L"");
                }
                states.push_back(std::move(state));
            }

            return states;
        };
        const auto formatEditableStates = [&](const std::vector<UiaValuePatternState>& states) noexcept
        {
            if (states.empty())
            {
                return std::wstring(L"<none>");
            }

            std::wstring details;
            for (size_t index = 0; index < states.size(); ++index)
            {
                if (! details.empty())
                {
                    details += L"; ";
                }
                details += std::format(L"[{0}] name='{1}' value='{2}'", index, states[index].name, states[index].value);
            }
            return details;
        };
        const auto waitForEditFocus = [&](HWND expectedDialog) noexcept
        {
            PluginConfigurationDialogDebugSnapshot snapshot{};
            return waitForSnapshot(expectedDialog, [](const PluginConfigurationDialogDebugSnapshot& value) noexcept {
                return value.focusKind == PluginConfigurationDialogDebugFocusKind::Edit && ! value.focusLabel.empty() && value.visibleDxFormInputHostCount > 0u;
            }, snapshot);
        };
        const auto collectFocusedEditableState = [&]() noexcept -> std::optional<UiaValuePatternState>
        {
            HWND focusedHost = nullptr;
            if (! DebugGetPluginConfigurationDialogFocusedHost(focusedHost) || ! focusedHost || IsWindow(focusedHost) == FALSE)
            {
                return std::nullopt;
            }

            const auto valueState = CollectWindowRootOrDescendantValuePatternState(focusedHost, UIA_EditControlTypeId);
            if (! valueState.has_value() || valueState->isReadOnly)
            {
                return std::nullopt;
            }

            return valueState;
        };
        const auto setFocusedEditValue = [&](std::wstring_view value) noexcept
        {
            HWND focusedHost = nullptr;
            if (! DebugGetPluginConfigurationDialogFocusedHost(focusedHost) || ! focusedHost || IsWindow(focusedHost) == FALSE)
            {
                return false;
            }

            return SetWindowRootOrDescendantValue(focusedHost, UIA_EditControlTypeId, value);
        };

        const std::wstring browseButtonText = LoadStringResource(nullptr, IDS_PREFS_PLUGINS_DETAILS_CONFIG_BROWSE_ELLIPSIS);
        if (! browseButtonText.empty())
        {
            const auto hasVisibleButtonNamed = [&](HWND targetDialog, std::wstring_view expectedName) noexcept
            {
                for (const auto& element : FindMatchingVisibleDescendantElements(targetDialog, UIA_ButtonControlTypeId))
                {
                    wil::unique_bstr currentName;
                    if (FAILED(element->get_CurrentName(&currentName)))
                    {
                        continue;
                    }

                    if ((currentName.get() ? std::wstring_view(currentName.get()) : std::wstring_view{}) == expectedName)
                    {
                        return true;
                    }
                }

                return false;
            };

            workerResult.browseButtonAvailable = hasVisibleButtonNamed(dialog, browseButtonText);
            traceStep(L"browse cancel begin");
            if (workerResult.browseButtonAvailable)
            {
                const auto editableStatesBeforeBrowse        = collectVisibleEditableStates(dialog);
                const auto anyVisibleEditDiffersFromBaseline = [&]() noexcept
                {
                    const auto states = collectVisibleEditableStates(dialog);
                    if (states.size() != editableStatesBeforeBrowse.size())
                    {
                        return true;
                    }

                    for (const auto& baselineState : editableStatesBeforeBrowse)
                    {
                        const auto match = std::find_if(
                            states.begin(), states.end(), [&](const UiaValuePatternState& candidate) noexcept { return candidate.name == baselineState.name; });
                        if (match == states.end() || match->value != baselineState.value)
                        {
                            return true;
                        }
                    }

                    return false;
                };
                const auto waitForVisibleEditsUnchanged = [&]() noexcept
                {
                    const auto deadline = std::chrono::steady_clock::now() + SelfTest::Scale(1500ms);
                    while (std::chrono::steady_clock::now() < deadline)
                    {
                        PumpPendingMessages();
                        if (! anyVisibleEditDiffersFromBaseline())
                        {
                            return true;
                        }
                        std::this_thread::sleep_for(20ms);
                    }

                    return ! anyVisibleEditDiffersFromBaseline();
                };

                const auto waitForAnyEditValue = [&](std::wstring_view expectedValue, std::wstring* matchedName, std::wstring* matchedPreviousValue) noexcept
                {
                    const auto matchesVisibleEditValue = [&](std::wstring_view candidateValue) noexcept
                    {
                        const auto states = collectVisibleEditableStates(dialog);
                        for (const auto& state : states)
                        {
                            if (state.value == candidateValue)
                            {
                                if (matchedName)
                                {
                                    *matchedName = state.name;
                                }
                                if (matchedPreviousValue)
                                {
                                    matchedPreviousValue->clear();
                                    for (const auto& baselineState : editableStatesBeforeBrowse)
                                    {
                                        if (baselineState.name == state.name)
                                        {
                                            *matchedPreviousValue = baselineState.value;
                                            break;
                                        }
                                    }
                                }
                                return true;
                            }
                        }

                        return false;
                    };

                    const auto deadline = std::chrono::steady_clock::now() + SelfTest::Scale(3000ms);
                    while (std::chrono::steady_clock::now() < deadline)
                    {
                        PumpPendingMessages();
                        if (matchesVisibleEditValue(expectedValue))
                        {
                            return true;
                        }
                        std::this_thread::sleep_for(20ms);
                    }

                    return matchesVisibleEditValue(expectedValue);
                };

                workerResult.sawBrowseButton =
                    DebugCancelPluginConfigurationNextBrowse() && InvokeVisibleDescendantByName(dialog, UIA_ButtonControlTypeId, browseButtonText);
                if (workerResult.sawBrowseButton)
                {
                    workerResult.browseCancelPreservedVisibleEdit = waitForVisibleEditsUnchanged();
                    if (workerResult.browseCancelPreservedVisibleEdit && DebugSetPluginConfigurationNextBrowsePath(browseTarget.native()))
                    {
                        traceStep(L"browse apply begin");
                        workerResult.sawBrowseButton = InvokeVisibleDescendantByName(dialog, UIA_ButtonControlTypeId, browseButtonText);
                        if (workerResult.sawBrowseButton)
                        {
                            workerResult.browseUpdatedVisibleEdit = waitForAnyEditValue(browseTarget.native(), &browseFieldName, &browseFieldBaselineValue);
                        }
                    }
                }
            }
            traceStep(L"browse phase done");
        }

        const auto editableStates    = collectVisibleEditableStates(dialog);
        workerResult.editDiagnostics = std::format(L"visibleEditableCount={}; candidates={}", editableStates.size(), formatEditableStates(editableStates));
        if (DebugFocusPluginConfigurationDialogFirstInput())
        {
            traceStep(L"edit phase begin");
            workerResult.editDiagnostics += L" | focus-requested=true";
            if (! waitForEditFocus(dialog))
            {
                workerResult.editDiagnostics += L" | focus-settle=false";
            }

            const auto waitForTrackedEditValue = [&](std::wstring_view expectedValue) noexcept
            {
                const auto deadline = std::chrono::steady_clock::now() + SelfTest::Scale(3000ms);
                while (std::chrono::steady_clock::now() < deadline)
                {
                    PumpPendingMessages();
                    const auto valueState = collectFocusedEditableState();
                    if (valueState.has_value() && valueState->value == expectedValue)
                    {
                        return true;
                    }
                    std::this_thread::sleep_for(20ms);
                }

                const auto valueState = collectFocusedEditableState();
                return valueState.has_value() && valueState->value == expectedValue;
            };
            if (const auto focusedEditState = collectFocusedEditableState(); focusedEditState.has_value())
            {
                editFieldName    = focusedEditState->name;
                initialEditValue = focusedEditState->value;
                editedValue      = (initialEditValue == L"selftest-plugin-config") ? L"selftest-plugin-config-2" : L"selftest-plugin-config";
                workerResult.editDiagnostics += std::format(L" | focused name='{0}' initial='{1}'", focusedEditState->name, focusedEditState->value);
                if (! setFocusedEditValue(editedValue))
                {
                    workerResult.editDiagnostics += std::format(L" set-failed target='{0}'", editedValue);
                }
                else
                {
                    workerResult.mutatedEdit = waitForTrackedEditValue(editedValue);
                    workerResult.editDiagnostics +=
                        std::format(L" set-ok target='{0}' mutated={1}", editedValue, workerResult.mutatedEdit ? L"true" : L"false");
                    if (! workerResult.mutatedEdit)
                    {
                        if (const auto observedFocusedState = collectFocusedEditableState(); observedFocusedState.has_value())
                        {
                            workerResult.editDiagnostics +=
                                std::format(L" observedFocused name='{0}' value='{1}'", observedFocusedState->name, observedFocusedState->value);
                        }
                        const auto observedStates = collectVisibleEditableStates(dialog);
                        workerResult.editDiagnostics += std::format(L" observedVisible={}", formatEditableStates(observedStates));
                    }

                    if (workerResult.mutatedEdit && setFocusedEditValue(initialEditValue))
                    {
                        workerResult.restoredEdit = waitForTrackedEditValue(initialEditValue);
                        workerResult.editDiagnostics += std::format(L" restored={0}", workerResult.restoredEdit ? L"true" : L"false");
                    }
                }
            }
            else
            {
                workerResult.editDiagnostics += L" | focused-edit=<none>";
            }
            traceStep(L"edit phase done");
        }
        else
        {
            workerResult.editDiagnostics += L" | focus-requested=false";
        }

        const auto initialToggleState = CollectVisibleDescendantTogglePatternState(dialog);
        if (initialToggleState.has_value() && ! initialToggleState->name.empty())
        {
            traceStep(L"toggle phase begin");
            toggleName                           = initialToggleState->name;
            initialToggleValue                   = initialToggleState->toggleState;
            const ToggleState flippedToggleValue = (*initialToggleValue == ToggleState_On) ? ToggleState_Off : ToggleState_On;

            if (ToggleVisibleDescendantByName(dialog, toggleName))
            {
                const auto waitForToggleState = [&](const ToggleState expectedState) noexcept
                {
                    const auto deadline = std::chrono::steady_clock::now() + SelfTest::Scale(3000ms);
                    while (std::chrono::steady_clock::now() < deadline)
                    {
                        PumpPendingMessages();
                        const auto toggleState = CollectVisibleDescendantTogglePatternState(dialog);
                        if (toggleState.has_value() && toggleState->toggleState == expectedState)
                        {
                            return true;
                        }
                        std::this_thread::sleep_for(20ms);
                    }

                    const auto toggleState = CollectVisibleDescendantTogglePatternState(dialog);
                    return toggleState.has_value() && toggleState->toggleState == expectedState;
                };

                if (waitForToggleState(flippedToggleValue))
                {
                    const auto flippedToggleState = CollectVisibleDescendantTogglePatternState(dialog);
                    workerResult.toggleLabelChanged =
                        flippedToggleState.has_value() && ! flippedToggleState->name.empty() && flippedToggleState->name != toggleName;
                    if (flippedToggleState.has_value() && ! flippedToggleState->name.empty() && ToggleVisibleDescendantByName(dialog, flippedToggleState->name))
                    {
                        workerResult.restoredToggle      = waitForToggleState(*initialToggleValue);
                        const auto restoredToggleState   = CollectVisibleDescendantTogglePatternState(dialog);
                        workerResult.toggleLabelRestored = restoredToggleState.has_value() && restoredToggleState->name == toggleName;
                    }
                }
            }
            traceStep(L"toggle phase done");
        }

        traceStep(L"final snapshot begin");
        workerResult.capturedFinalSnapshot = waitForSnapshot(dialog,
                                                             [](const PluginConfigurationDialogDebugSnapshot& value) noexcept
        {
            return value.usesDxUiCommandButtons && value.usesDxUiFormSurface && value.usesDxUiFormStatics && value.usesDxUiFormInputs &&
                   value.visibleLegacyCommandButtonCount == 0u && value.visibleLegacyFormControlCount == 0u && value.visibleDxCommandButtonHostCount == 2u &&
                   value.visibleDxFormInputHostCount > 0u;
        },
                                                             workerResult.finalSnapshot);
        traceStep(L"final snapshot done");

        const std::wstring cancelButtonText = LoadStringResource(nullptr, IDS_BTN_CANCEL);
        traceStep(L"cancel invoke begin");
        workerResult.invokedCancel = ! cancelButtonText.empty() && InvokeVisibleDescendantByName(dialog, UIA_ButtonControlTypeId, cancelButtonText);
        if (workerResult.invokedCancel)
        {
            workerResult.closedAfterCancel = WaitForWindowClosed(dialog, SelfTest::Scale(3000ms));
        }
        if (! workerResult.closedAfterCancel)
        {
            traceStep(L"cancel invoke failed");
            static_cast<void>(DebugCancelPluginConfigurationDialog());
            return;
        }
        traceStep(L"cancel invoke done");

        traceStep(L"wait reopened dialog");
        const HWND reopenedDialog      = WaitForWindow([] noexcept { return GetPluginConfigurationDialogHandle(); }, SelfTest::Scale(5000ms));
        workerResult.sawReopenedDialog = reopenedDialog != nullptr && IsWindow(reopenedDialog) != FALSE;
        if (! workerResult.sawReopenedDialog)
        {
            return;
        }

        traceStep(L"reopened dialog ready");
        workerResult.reopenedOwnedByMainWindow = IsOwnedBy(reopenedDialog, mainWindow);
        workerResult.capturedReopenedSnapshot  = waitForSnapshot(reopenedDialog,
                                                                 [](const PluginConfigurationDialogDebugSnapshot& value) noexcept
        {
            return value.usesDxUiCommandButtons && value.usesDxUiFormSurface && value.usesDxUiFormStatics && value.usesDxUiFormInputs &&
                   value.visibleLegacyCommandButtonCount == 0u && value.visibleLegacyFormControlCount == 0u && value.visibleDxCommandButtonHostCount == 2u &&
                   value.visibleDxFormInputHostCount > 0u;
        },
                                                                workerResult.reopenedSnapshot);
        traceStep(L"reopened snapshot done");

        bool refocusedReopenedEdit = false;
        if (DebugFocusPluginConfigurationDialogFirstInput())
        {
            refocusedReopenedEdit = waitForEditFocus(reopenedDialog);
        }
        workerResult.editDiagnostics += std::format(L" | reopened-focus={}", refocusedReopenedEdit ? L"true" : L"false");
        const auto waitForReopenedEditValue = [&](std::wstring_view expectedValue) noexcept
        {
            const auto deadline = std::chrono::steady_clock::now() + SelfTest::Scale(3000ms);
            while (std::chrono::steady_clock::now() < deadline)
            {
                PumpPendingMessages();
                const auto valueState = collectFocusedEditableState();
                if (valueState.has_value() && valueState->value == expectedValue)
                {
                    return true;
                }
                std::this_thread::sleep_for(20ms);
            }

            const auto valueState = collectFocusedEditableState();
            return valueState.has_value() && valueState->value == expectedValue;
        };
        const auto waitForReopenedNamedEditValue = [&](std::wstring_view expectedName, std::wstring_view expectedValue) noexcept
        {
            const auto deadline = std::chrono::steady_clock::now() + SelfTest::Scale(3000ms);
            while (std::chrono::steady_clock::now() < deadline)
            {
                PumpPendingMessages();
                const auto valueState = CollectVisibleDescendantValuePatternStateByName(reopenedDialog, UIA_EditControlTypeId, expectedName);
                if (valueState.has_value() && valueState->value == expectedValue)
                {
                    return true;
                }
                std::this_thread::sleep_for(20ms);
            }

            const auto valueState = CollectVisibleDescendantValuePatternStateByName(reopenedDialog, UIA_EditControlTypeId, expectedName);
            return valueState.has_value() && valueState->value == expectedValue;
        };
        if (! editedValue.empty())
        {
            traceStep(L"reopened edit restore begin");
            workerResult.reopenedEditRestored = waitForReopenedEditValue(initialEditValue);
        }
        if (! browseFieldName.empty())
        {
            traceStep(L"reopened browse restore begin");
            workerResult.browseRestoredAfterCancel = waitForReopenedNamedEditValue(browseFieldName, browseFieldBaselineValue);
        }
        if (! browseFieldName.empty() && ! browseButtonText.empty() && DebugSetPluginConfigurationNextBrowsePath(browseTarget.native()) &&
            InvokeVisibleDescendantByName(reopenedDialog, UIA_ButtonControlTypeId, browseButtonText))
        {
            traceStep(L"reopened browse roundtrip begin");
            if (waitForReopenedNamedEditValue(browseFieldName, browseTarget.native()) &&
                SetVisibleDescendantValueByName(reopenedDialog, UIA_EditControlTypeId, browseFieldName, browseFieldBaselineValue))
            {
                workerResult.reopenedBrowseRoundTrip = waitForReopenedNamedEditValue(browseFieldName, browseFieldBaselineValue);
            }
            traceStep(L"reopened browse roundtrip done");
        }

        const auto setReopenedTrackedEditValue = [&](std::wstring_view value) noexcept { return setFocusedEditValue(value); };

        if (! editedValue.empty() && setReopenedTrackedEditValue(editedValue))
        {
            traceStep(L"reopened edit roundtrip begin");
            if (waitForReopenedEditValue(editedValue) && setReopenedTrackedEditValue(initialEditValue))
            {
                workerResult.reopenedEditRoundTrip = waitForReopenedEditValue(initialEditValue);
            }
            traceStep(L"reopened edit roundtrip done");
        }

        if (! toggleName.empty() && initialToggleValue.has_value())
        {
            traceStep(L"reopened toggle phase begin");
            const auto waitForReopenedToggleState = [&](const ToggleState expectedState) noexcept
            {
                const auto deadline = std::chrono::steady_clock::now() + SelfTest::Scale(3000ms);
                while (std::chrono::steady_clock::now() < deadline)
                {
                    PumpPendingMessages();
                    const auto toggleState = CollectVisibleDescendantTogglePatternState(reopenedDialog);
                    if (toggleState.has_value() && toggleState->toggleState == expectedState)
                    {
                        return true;
                    }
                    std::this_thread::sleep_for(20ms);
                }

                const auto toggleState = CollectVisibleDescendantTogglePatternState(reopenedDialog);
                return toggleState.has_value() && toggleState->toggleState == expectedState;
            };
            workerResult.reopenedToggleRestored      = waitForReopenedToggleState(*initialToggleValue);
            const auto reopenedBaselineToggleState   = CollectVisibleDescendantTogglePatternState(reopenedDialog);
            workerResult.reopenedToggleLabelRestored = reopenedBaselineToggleState.has_value() && reopenedBaselineToggleState->name == toggleName;

            const ToggleState flippedToggleValue = (*initialToggleValue == ToggleState_On) ? ToggleState_Off : ToggleState_On;
            if (workerResult.reopenedToggleRestored && ToggleVisibleDescendantByName(reopenedDialog, toggleName))
            {
                const bool sawReopenedFlippedToggleState = waitForReopenedToggleState(flippedToggleValue);
                const auto reopenedFlippedToggleState    = CollectVisibleDescendantTogglePatternState(reopenedDialog);
                if (sawReopenedFlippedToggleState && reopenedFlippedToggleState.has_value() && ! reopenedFlippedToggleState->name.empty() &&
                    ToggleVisibleDescendantByName(reopenedDialog, reopenedFlippedToggleState->name))
                {
                    workerResult.reopenedToggleRoundTrip      = waitForReopenedToggleState(*initialToggleValue);
                    const auto reopenedRestoredToggleState    = CollectVisibleDescendantTogglePatternState(reopenedDialog);
                    workerResult.reopenedToggleLabelRoundTrip = reopenedFlippedToggleState.has_value() && ! reopenedFlippedToggleState->name.empty() &&
                                                                reopenedFlippedToggleState->name != toggleName && reopenedRestoredToggleState.has_value() &&
                                                                reopenedRestoredToggleState->name == toggleName;
                }
            }
            traceStep(L"reopened toggle phase done");
        }

        const std::wstring okButtonText = LoadStringResource(nullptr, IDS_BTN_OK);
        traceStep(L"ok invoke begin");
        workerResult.invokedOk = ! okButtonText.empty() && InvokeVisibleDescendantByName(reopenedDialog, UIA_ButtonControlTypeId, okButtonText);
        if (workerResult.invokedOk)
        {
            workerResult.closedAfterInvoke = WaitForWindowClosed(reopenedDialog, SelfTest::Scale(3000ms));
        }

        if (! workerResult.closedAfterInvoke)
        {
            traceStep(L"ok invoke failed");
            static_cast<void>(DebugCancelPluginConfigurationDialog());
            return;
        }

        traceStep(L"ok invoke done");
    });

    Common::Settings::Settings baselineSettings = g_settings;
    Common::Settings::Settings workingSettings  = baselineSettings;
    const AppTheme theme                        = ResolveAppTheme(ThemeMode::Dark, L"plugin-config-interaction-selftest");
    const HRESULT cancelHr =
        EditPluginConfigurationDialog(mainWindow, PluginType::FileSystem, kBuiltinS3FileSystemId, entry->name, baselineSettings, workingSettings, theme);
    const HRESULT okHr =
        EditPluginConfigurationDialog(mainWindow, PluginType::FileSystem, kBuiltinS3FileSystemId, entry->name, baselineSettings, workingSettings, theme);
    worker.join();

    state.Require(workerResult.sawDialog, L"Plugin configuration dialog did not open for live interaction validation.");
    state.Require(workerResult.ownedByMainWindow, L"Plugin configuration dialog should be owned by the main window during live interaction validation.");
    state.Require(workerResult.capturedBaselineSnapshot, L"Failed to capture plugin configuration dialog baseline snapshot for live interaction validation.");
    if (workerResult.browseButtonAvailable)
    {
        state.Require(workerResult.sawBrowseButton,
                      L"Plugin configuration dialog visible DX Browse... action did not expose live UIA InvokePattern interaction.");
        state.Require(workerResult.browseCancelPreservedVisibleEdit,
                      L"Plugin configuration dialog visible DX Browse... action did not preserve visible editable-field state after the canceled browse path.");
        state.Require(workerResult.browseUpdatedVisibleEdit,
                      L"Plugin configuration dialog visible DX Browse... action did not update any visible editable field with the browsed folder path.");
        state.Require(workerResult.browseRestoredAfterCancel,
                      L"Plugin configuration dialog did not restore the browsed field to its baseline value after live DX Cancel interaction.");
    }
    state.Require(workerResult.restoredToggle, L"Plugin configuration dialog visible DX toggle did not round-trip through live UIA TogglePattern interaction.");
    state.Require(workerResult.toggleLabelChanged,
                  L"Plugin configuration dialog visible DX toggle did not refresh its visible label after live UIA TogglePattern interaction.");
    state.Require(workerResult.toggleLabelRestored,
                  L"Plugin configuration dialog visible DX toggle did not restore its baseline visible label after the live toggle round-trip.");
    state.Require(workerResult.mutatedEdit,
                  std::format(L"Plugin configuration dialog visible DX edit did not expose live UIA ValuePattern mutation interaction. {}",
                              workerResult.editDiagnostics));
    state.Require(workerResult.capturedFinalSnapshot, L"Failed to capture plugin configuration dialog final snapshot after live interaction validation.");
    state.Require(workerResult.finalSnapshot.usesDxUiCommandButtons && workerResult.finalSnapshot.usesDxUiFormSurface &&
                      workerResult.finalSnapshot.usesDxUiFormStatics && workerResult.finalSnapshot.usesDxUiFormInputs,
                  L"Plugin configuration dialog lost the active DxUi shell/form surface during live interaction validation.");
    state.Require(workerResult.finalSnapshot.visibleLegacyCommandButtonCount == 0u && workerResult.finalSnapshot.visibleLegacyFormControlCount == 0u,
                  L"Plugin configuration dialog exposed visible legacy controls during live interaction validation.");
    state.Require(workerResult.invokedCancel, L"Plugin configuration dialog visible DX Cancel action did not expose live UIA InvokePattern interaction.");
    state.Require(workerResult.closedAfterCancel,
                  L"Plugin configuration dialog did not close after live UIA InvokePattern interaction on the visible DX Cancel action.");
    state.Require(workerResult.sawReopenedDialog, L"Plugin configuration dialog did not reopen after the live DX Cancel-action validation pass.");
    state.Require(workerResult.reopenedOwnedByMainWindow,
                  L"Plugin configuration dialog reopened without the expected main-window ownership during live interaction validation.");
    state.Require(workerResult.capturedReopenedSnapshot,
                  L"Failed to capture plugin configuration dialog reopened snapshot after the live DX Cancel-action validation pass.");
    state.Require(workerResult.reopenedEditRestored,
                  L"Plugin configuration dialog reopened without restoring the visible DX edit to its baseline value after live DX Cancel interaction.");
    state.Require(workerResult.reopenedToggleRestored,
                  L"Plugin configuration dialog reopened without restoring the visible DX toggle to its baseline state after live DX Cancel interaction.");
    state.Require(workerResult.reopenedToggleLabelRestored,
                  L"Plugin configuration dialog reopened without restoring the visible DX toggle label to its baseline text after live DX Cancel interaction.");
    if (workerResult.browseButtonAvailable)
    {
        state.Require(
            workerResult.reopenedBrowseRoundTrip,
            L"Plugin configuration dialog reopened without rerunning the visible DX Browse... interaction back to its baseline field state before confirm.");
    }
    state.Require(workerResult.reopenedEditRoundTrip, L"Plugin configuration dialog reopened without rerunning the visible DX edit round-trip before confirm.");
    state.Require(workerResult.reopenedToggleRoundTrip,
                  L"Plugin configuration dialog reopened without rerunning the visible DX toggle round-trip before confirm.");
    state.Require(workerResult.reopenedToggleLabelRoundTrip,
                  L"Plugin configuration dialog reopened without rerunning the visible DX toggle label change-and-restore path before confirm.");
    state.Require(workerResult.invokedOk, L"Plugin configuration dialog visible DX OK action did not expose live UIA InvokePattern interaction.");
    state.Require(workerResult.closedAfterInvoke,
                  L"Plugin configuration dialog did not close after live UIA InvokePattern interaction on the visible DX OK action.");
    state.Require(! workerResult.blockedStep.empty(), L"Plugin configuration live interaction selftest did not report any progress breadcrumbs.");
    state.Require(cancelHr == S_FALSE,
                  std::format(L"EditPluginConfigurationDialog returned unexpected HRESULT 0x{:08X} for the cancel pass during live interaction validation.",
                              static_cast<uint32_t>(cancelHr)));
    state.Require(okHr == S_OK,
                  std::format(L"EditPluginConfigurationDialog returned unexpected HRESULT 0x{:08X} for the confirm pass during live interaction validation.",
                              static_cast<uint32_t>(okHr)));
    return state.failure.empty();
}

[[nodiscard]] bool TestPluginConfigurationDialogTabTraversalLiveDxInteraction(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    constexpr std::wstring_view kBuiltinS3FileSystemId = L"builtin/file-system-s3";
    const FileSystemPluginManager::PluginEntry* entry  = FindFileSystemPluginById(kBuiltinS3FileSystemId);
    state.Require(entry != nullptr, L"builtin/file-system-s3 plugin entry unavailable for plugin configuration tab-traversal self-test.");
    if (! entry)
    {
        return false;
    }

    struct WorkerResult
    {
        bool sawDialog                = false;
        bool ownedByMainWindow        = false;
        bool capturedBaselineSnapshot = false;
        bool focusedFirstInput        = false;
        bool completedTraversal       = false;
        bool cancelled                = false;
        std::wstring failedTraversalStep;
        std::optional<UiaDescendantPatternStats> patternStats;
        PluginConfigurationDialogDebugSnapshot baselineSnapshot{};
        PluginConfigurationDialogDebugSnapshot traversalSnapshot{};
    } workerResult{};

    const auto baselineSnapshotReady = [](const PluginConfigurationDialogDebugSnapshot& value) noexcept
    {
        return value.usesDxUiCommandButtons && value.usesDxUiFormSurface && value.usesDxUiFormStatics && value.usesDxUiFormInputs &&
               value.visibleLegacyCommandButtonCount == 0u && value.visibleLegacyFormControlCount == 0u && value.visibleDxCommandButtonHostCount == 2u &&
               value.visibleDxFormInputHostCount > 0u && value.panelScrollPosY == 0;
    };

    std::jthread worker([&](std::stop_token) noexcept
    {
        const HWND dialog      = WaitForWindow([] noexcept { return GetPluginConfigurationDialogHandle(); }, SelfTest::Scale(5000ms));
        workerResult.sawDialog = dialog != nullptr && IsWindow(dialog) != FALSE;
        if (! workerResult.sawDialog)
        {
            return;
        }

        workerResult.ownedByMainWindow = IsOwnedBy(dialog, mainWindow);

        const auto waitForSnapshot = [&](const auto& predicate, PluginConfigurationDialogDebugSnapshot& outSnapshot) noexcept
        {
            const auto deadline = std::chrono::steady_clock::now() + SelfTest::Scale(3000ms);
            while (std::chrono::steady_clock::now() < deadline)
            {
                PumpPendingMessages();
                outSnapshot = {};
                if (DebugGetPluginConfigurationDialogSnapshot(outSnapshot) && predicate(outSnapshot))
                {
                    return true;
                }
                std::this_thread::sleep_for(20ms);
            }

            outSnapshot = {};
            return DebugGetPluginConfigurationDialogSnapshot(outSnapshot) && predicate(outSnapshot);
        };

        workerResult.capturedBaselineSnapshot = waitForSnapshot(baselineSnapshotReady, workerResult.baselineSnapshot);
        if (! workerResult.capturedBaselineSnapshot && baselineSnapshotReady(workerResult.baselineSnapshot))
        {
            workerResult.capturedBaselineSnapshot = true;
        }
        if (! workerResult.capturedBaselineSnapshot)
        {
            static_cast<void>(DebugCancelPluginConfigurationDialog());
            return;
        }

        workerResult.patternStats = CollectVisibleUiaDescendantPatternStats(dialog);
        if (workerResult.baselineSnapshot.focusKind == PluginConfigurationDialogDebugFocusKind::Edit &&
            workerResult.baselineSnapshot.focusLabel == L"Default region")
        {
            workerResult.focusedFirstInput = true;
            workerResult.traversalSnapshot = workerResult.baselineSnapshot;
        }
        else if (! DebugFocusPluginConfigurationDialogFirstInput())
        {
            SelfTest::AppendSuiteTrace(SelfTest::SelfTestSuite::Commands,
                                       std::format(L"plugin-config tab focus command failed: baselineFocusKind={} baselineFocusLabel='{}' dxButtons={} "
                                                   L"dxStatics={} dxInputs={} legacyButtons={} legacyControls={} scrollY={}",
                                                   static_cast<int>(workerResult.baselineSnapshot.focusKind),
                                                   workerResult.baselineSnapshot.focusLabel,
                                                   workerResult.baselineSnapshot.visibleDxCommandButtonHostCount,
                                                   workerResult.baselineSnapshot.visibleDxFormStaticHostCount,
                                                   workerResult.baselineSnapshot.visibleDxFormInputHostCount,
                                                   workerResult.baselineSnapshot.visibleLegacyCommandButtonCount,
                                                   workerResult.baselineSnapshot.visibleLegacyFormControlCount,
                                                   workerResult.baselineSnapshot.panelScrollPosY));
            static_cast<void>(DebugCancelPluginConfigurationDialog());
            return;
        }
        else
        {
            workerResult.focusedFirstInput = waitForSnapshot(
                [&](const PluginConfigurationDialogDebugSnapshot& value) noexcept
            {
                return value.focusKind == PluginConfigurationDialogDebugFocusKind::Edit && value.focusLabel == L"Default region" &&
                       value.visibleDxCommandButtonHostCount == workerResult.baselineSnapshot.visibleDxCommandButtonHostCount &&
                       value.visibleDxFormStaticHostCount == workerResult.baselineSnapshot.visibleDxFormStaticHostCount &&
                       value.visibleDxFormInputHostCount == workerResult.baselineSnapshot.visibleDxFormInputHostCount && value.panelScrollPosY == 0;
            },
                workerResult.traversalSnapshot);
            if (! workerResult.focusedFirstInput)
            {
                SelfTest::AppendSuiteTrace(SelfTest::SelfTestSuite::Commands,
                                           std::format(L"plugin-config tab focus miss: focusKind={} focusLabel='{}' dxButtons={} dxStatics={} dxInputs={} "
                                                       L"legacyButtons={} legacyControls={} scrollY={}",
                                                       static_cast<int>(workerResult.traversalSnapshot.focusKind),
                                                       workerResult.traversalSnapshot.focusLabel,
                                                       workerResult.traversalSnapshot.visibleDxCommandButtonHostCount,
                                                       workerResult.traversalSnapshot.visibleDxFormStaticHostCount,
                                                       workerResult.traversalSnapshot.visibleDxFormInputHostCount,
                                                       workerResult.traversalSnapshot.visibleLegacyCommandButtonCount,
                                                       workerResult.traversalSnapshot.visibleLegacyFormControlCount,
                                                       workerResult.traversalSnapshot.panelScrollPosY));
                static_cast<void>(DebugCancelPluginConfigurationDialog());
                return;
            }
        }

        const auto sendTab = [&](const bool reverse,
                                 const PluginConfigurationDialogDebugFocusKind expectedKind,
                                 std::wstring_view expectedLabel,
                                 std::wstring_view label) noexcept
        {
            const auto advanceAndWait = [&]() noexcept
            {
                const bool advanced = DebugAdvancePluginConfigurationDialogTab(reverse);
                PumpPendingMessages();
                if (! advanced)
                {
                    workerResult.failedTraversalStep.assign(label);
                    SelfTest::AppendSuiteTrace(SelfTest::SelfTestSuite::Commands,
                                               std::format(L"plugin-config tab advance command failed: step='{}' reverse={}", label, reverse ? 1 : 0));
                    return false;
                }

                return waitForSnapshot(
                    [&](const PluginConfigurationDialogDebugSnapshot& value) noexcept
                {
                    return value.focusKind == expectedKind && value.focusLabel == expectedLabel && value.usesDxUiCommandButtons && value.usesDxUiFormSurface &&
                           value.usesDxUiFormStatics && value.usesDxUiFormInputs && value.visibleDxCommandButtonHostCount == 2u &&
                           value.visibleDxFormStaticHostCount > 0u && value.visibleDxFormInputHostCount > 0u && value.visibleLegacyCommandButtonCount == 0u &&
                           value.visibleLegacyFormControlCount == 0u && value.visibleLegacyFormStaticCount == 0u && value.panelScrollPosY == 0;
                },
                    workerResult.traversalSnapshot);
            };

            bool reached = advanceAndWait();
            if (! reached && ! reverse && label == L"default endpoint override edit" &&
                workerResult.traversalSnapshot.focusKind == PluginConfigurationDialogDebugFocusKind::None && workerResult.traversalSnapshot.focusLabel.empty())
            {
                const bool recoveredFirstInputFocus = DebugFocusPluginConfigurationDialogFirstInput();
                PumpPendingMessages();
                if (recoveredFirstInputFocus)
                {
                    static_cast<void>(waitForSnapshot(
                        [&](const PluginConfigurationDialogDebugSnapshot& value) noexcept
                    {
                        return value.focusKind == PluginConfigurationDialogDebugFocusKind::Edit && value.focusLabel == L"Default region" &&
                               value.visibleDxCommandButtonHostCount == workerResult.baselineSnapshot.visibleDxCommandButtonHostCount &&
                               value.visibleDxFormStaticHostCount >= workerResult.baselineSnapshot.visibleDxFormStaticHostCount &&
                               value.visibleDxFormInputHostCount >= workerResult.baselineSnapshot.visibleDxFormInputHostCount && value.panelScrollPosY == 0;
                    },
                        workerResult.traversalSnapshot));
                    reached = advanceAndWait();
                }
            }

            if (! reached)
            {
                PumpPendingMessages();
                reached = advanceAndWait();
            }

            if (! reached)
            {
                workerResult.failedTraversalStep.assign(label);
                SelfTest::AppendSuiteTrace(SelfTest::SelfTestSuite::Commands,
                                           std::format(L"plugin-config tab step miss: step='{}' expectedKind={} expectedLabel='{}' labelMatches={} "
                                                       L"actualFocusKind={} actualFocusLabel='{}' usesButtons={} usesSurface={} usesStatics={} usesInputs={} "
                                                       L"dxButtons={} dxStatics={} dxInputs={} legacyButtons={} legacyControls={} legacyStatics={} scrollY={}",
                                                       label,
                                                       static_cast<int>(expectedKind),
                                                       expectedLabel,
                                                       workerResult.traversalSnapshot.focusLabel == expectedLabel ? 1 : 0,
                                                       static_cast<int>(workerResult.traversalSnapshot.focusKind),
                                                       workerResult.traversalSnapshot.focusLabel,
                                                       workerResult.traversalSnapshot.usesDxUiCommandButtons ? 1 : 0,
                                                       workerResult.traversalSnapshot.usesDxUiFormSurface ? 1 : 0,
                                                       workerResult.traversalSnapshot.usesDxUiFormStatics ? 1 : 0,
                                                       workerResult.traversalSnapshot.usesDxUiFormInputs ? 1 : 0,
                                                       workerResult.traversalSnapshot.visibleDxCommandButtonHostCount,
                                                       workerResult.traversalSnapshot.visibleDxFormStaticHostCount,
                                                       workerResult.traversalSnapshot.visibleDxFormInputHostCount,
                                                       workerResult.traversalSnapshot.visibleLegacyCommandButtonCount,
                                                       workerResult.traversalSnapshot.visibleLegacyFormControlCount,
                                                       workerResult.traversalSnapshot.visibleLegacyFormStaticCount,
                                                       workerResult.traversalSnapshot.panelScrollPosY));
            }
            return reached;
        };

        if (! sendTab(false, PluginConfigurationDialogDebugFocusKind::Edit, L"Default endpoint override", L"default endpoint override edit") ||
            ! sendTab(false, PluginConfigurationDialogDebugFocusKind::Toggle, L"Use HTTPS", L"Use HTTPS toggle") ||
            ! sendTab(false, PluginConfigurationDialogDebugFocusKind::Toggle, L"Verify TLS certificate", L"Verify TLS certificate toggle") ||
            ! sendTab(false, PluginConfigurationDialogDebugFocusKind::Edit, L"Connect timeout (ms)", L"connect timeout edit") ||
            ! sendTab(false, PluginConfigurationDialogDebugFocusKind::Edit, L"Network stall timeout (ms)", L"request timeout edit") ||
            ! sendTab(false,
                      PluginConfigurationDialogDebugFocusKind::Toggle,
                      L"Use virtual-hosted style addressing",
                      L"use virtual-hosted style addressing toggle") ||
            ! sendTab(false, PluginConfigurationDialogDebugFocusKind::Edit, L"Max keys per request", L"max keys edit") ||
            ! sendTab(false, PluginConfigurationDialogDebugFocusKind::CommandButton, LoadStringResource(nullptr, IDS_BTN_OK), L"OK button") ||
            ! sendTab(false, PluginConfigurationDialogDebugFocusKind::CommandButton, LoadStringResource(nullptr, IDS_BTN_CANCEL), L"Cancel button") ||
            ! sendTab(true, PluginConfigurationDialogDebugFocusKind::CommandButton, LoadStringResource(nullptr, IDS_BTN_OK), L"reverse OK button") ||
            ! sendTab(true, PluginConfigurationDialogDebugFocusKind::Edit, L"Max keys per request", L"reverse max keys edit") ||
            ! sendTab(true,
                      PluginConfigurationDialogDebugFocusKind::Toggle,
                      L"Use virtual-hosted style addressing",
                      L"reverse use virtual-hosted style addressing toggle") ||
            ! sendTab(true, PluginConfigurationDialogDebugFocusKind::Edit, L"Network stall timeout (ms)", L"reverse request timeout edit") ||
            ! sendTab(true, PluginConfigurationDialogDebugFocusKind::Edit, L"Connect timeout (ms)", L"reverse connect timeout edit") ||
            ! sendTab(true, PluginConfigurationDialogDebugFocusKind::Toggle, L"Verify TLS certificate", L"reverse Verify TLS certificate toggle") ||
            ! sendTab(true, PluginConfigurationDialogDebugFocusKind::Toggle, L"Use HTTPS", L"reverse Use HTTPS toggle") ||
            ! sendTab(true, PluginConfigurationDialogDebugFocusKind::Edit, L"Default endpoint override", L"reverse default endpoint override edit") ||
            ! sendTab(true, PluginConfigurationDialogDebugFocusKind::Edit, L"Default region", L"reverse default region edit") ||
            ! sendTab(
                true, PluginConfigurationDialogDebugFocusKind::CommandButton, LoadStringResource(nullptr, IDS_BTN_CANCEL), L"reverse wrapped Cancel button") ||
            ! sendTab(false, PluginConfigurationDialogDebugFocusKind::Edit, L"Default region", L"wrapped default region edit"))
        {
            static_cast<void>(DebugCancelPluginConfigurationDialog());
            return;
        }

        workerResult.completedTraversal = true;
        workerResult.cancelled          = DebugCancelPluginConfigurationDialog();
    });

    Common::Settings::Settings baselineSettings = g_settings;
    Common::Settings::Settings workingSettings  = baselineSettings;
    const AppTheme theme                        = ResolveAppTheme(ThemeMode::Dark, L"plugin-config-tab-traversal-selftest");
    const HRESULT hr =
        EditPluginConfigurationDialog(mainWindow, PluginType::FileSystem, kBuiltinS3FileSystemId, entry->name, baselineSettings, workingSettings, theme);
    worker.join();

    state.Require(workerResult.sawDialog, L"Plugin configuration dialog did not open for tab-traversal validation.");
    state.Require(workerResult.ownedByMainWindow, L"Plugin configuration dialog should be owned by the main window during tab-traversal validation.");
    state.Require(workerResult.capturedBaselineSnapshot, L"Failed to capture plugin configuration dialog baseline snapshot for tab-traversal validation.");
    state.Require(workerResult.patternStats.has_value(),
                  L"Failed to collect UI Automation pattern statistics for plugin configuration tab-traversal validation.");
    if (workerResult.patternStats.has_value())
    {
        state.Require(workerResult.patternStats->editControlCount > 0u,
                      L"Plugin configuration dialog should expose visible DX edit descendants before tab traversal.");
        state.Require(workerResult.patternStats->togglePatternCount > 0u,
                      L"Plugin configuration dialog should expose visible DX toggle descendants before tab traversal.");
        state.Require(workerResult.patternStats->buttonControlCount > 0u,
                      L"Plugin configuration dialog should expose visible DX command buttons before tab traversal.");
        state.Require(workerResult.patternStats->invokePatternCount > 0u,
                      L"Plugin configuration dialog should expose live UIA InvokePattern support before tab traversal.");
    }
    state.Require(workerResult.focusedFirstInput, L"Plugin configuration dialog did not focus the first visible DX input before tab traversal.");
    state.Require(
        workerResult.completedTraversal,
        workerResult.failedTraversalStep.empty()
            ? L"Plugin configuration dialog tab traversal did not complete."
            : std::format(L"Plugin configuration dialog tab traversal did not complete; last focused label was '{}'.", workerResult.failedTraversalStep));
    state.Require(workerResult.cancelled, L"Plugin configuration dialog debug cancel command failed after tab-traversal validation.");
    state.Require(
        hr == S_FALSE,
        std::format(L"EditPluginConfigurationDialog returned unexpected HRESULT 0x{:08X} during tab-traversal validation.", static_cast<uint32_t>(hr)));
    return state.failure.empty();
}

[[nodiscard]] bool TestPluginConfigurationDialogEnterAndEscapeRouteDefaultCancel(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    constexpr std::wstring_view kBuiltinS3FileSystemId = L"builtin/file-system-s3";
    const FileSystemPluginManager::PluginEntry* entry  = FindFileSystemPluginById(kBuiltinS3FileSystemId);
    state.Require(entry != nullptr, L"builtin/file-system-s3 plugin entry unavailable for default/cancel keyboard validation.");
    if (! entry)
    {
        return false;
    }

    struct PassResult
    {
        bool sawDialog                = false;
        bool ownedByMainWindow        = false;
        bool capturedBaselineSnapshot = false;
        bool focusedFirstInput        = false;
        bool closedAfterKey           = false;
        bool closedAfterFallback      = false;
        std::optional<UiaDescendantPatternStats> patternStats;
        PluginConfigurationDialogDebugSnapshot baselineSnapshot{};
        PluginConfigurationDialogDebugSnapshot focusSnapshot{};
    };

    const auto runPass = [&](const bool accept, std::wstring_view label) noexcept
    {
        PassResult result{};

        std::jthread worker([&](std::stop_token) noexcept
        {
            const HWND dialog = WaitForWindow([] noexcept { return GetPluginConfigurationDialogHandle(); }, SelfTest::Scale(5000ms));
            result.sawDialog  = dialog != nullptr && IsWindow(dialog) != FALSE;
            if (! result.sawDialog)
            {
                return;
            }

            result.ownedByMainWindow = IsOwnedBy(dialog, mainWindow);

            const auto waitForSnapshot = [&](const auto& predicate, PluginConfigurationDialogDebugSnapshot& outSnapshot) noexcept
            {
                const auto deadline = std::chrono::steady_clock::now() + SelfTest::Scale(3000ms);
                while (std::chrono::steady_clock::now() < deadline)
                {
                    PumpPendingMessages();
                    outSnapshot = {};
                    if (DebugGetPluginConfigurationDialogSnapshot(outSnapshot) && predicate(outSnapshot))
                    {
                        return true;
                    }
                    std::this_thread::sleep_for(20ms);
                }

                outSnapshot = {};
                return DebugGetPluginConfigurationDialogSnapshot(outSnapshot) && predicate(outSnapshot);
            };

            result.capturedBaselineSnapshot = waitForSnapshot(
                [](const PluginConfigurationDialogDebugSnapshot& value) noexcept
            {
                return value.usesDxUiCommandButtons && value.usesDxUiFormSurface && value.usesDxUiFormStatics && value.usesDxUiFormInputs &&
                       value.visibleLegacyCommandButtonCount == 0u && value.visibleLegacyFormControlCount == 0u &&
                       value.visibleDxCommandButtonHostCount == 2u && value.visibleDxFormInputHostCount > 0u && value.panelScrollPosY == 0;
            },
                result.baselineSnapshot);
            if (! result.capturedBaselineSnapshot)
            {
                static_cast<void>(DebugCancelPluginConfigurationDialog());
                return;
            }

            result.patternStats                 = CollectVisibleUiaDescendantPatternStats(dialog);
            const bool requestedFirstInputFocus = DebugFocusPluginConfigurationDialogFirstInput();

            result.focusedFirstInput = waitForSnapshot(
                [&](const PluginConfigurationDialogDebugSnapshot& value) noexcept
            {
                const bool isDxInputFocus = value.focusKind == PluginConfigurationDialogDebugFocusKind::Edit ||
                                            value.focusKind == PluginConfigurationDialogDebugFocusKind::Toggle ||
                                            value.focusKind == PluginConfigurationDialogDebugFocusKind::Combo;
                return isDxInputFocus && ! value.focusLabel.empty() &&
                       value.visibleDxCommandButtonHostCount >= result.baselineSnapshot.visibleDxCommandButtonHostCount &&
                       value.visibleDxFormStaticHostCount >= result.baselineSnapshot.visibleDxFormStaticHostCount &&
                       value.visibleDxFormInputHostCount >= result.baselineSnapshot.visibleDxFormInputHostCount && value.panelScrollPosY == 0;
            },
                result.focusSnapshot);
            if (! result.focusedFirstInput)
            {
                if (! requestedFirstInputFocus)
                {
                    static_cast<void>(DebugCancelPluginConfigurationDialog());
                    return;
                }
                static_cast<void>(DebugCancelPluginConfigurationDialog());
                return;
            }

            HWND focusedHost = nullptr;
            if (! DebugGetPluginConfigurationDialogFocusedHost(focusedHost) || ! focusedHost || IsWindow(focusedHost) == FALSE)
            {
                static_cast<void>(DebugCancelPluginConfigurationDialog());
                return;
            }

            const WPARAM vk = accept ? VK_RETURN : VK_ESCAPE;
            SendMessageW(focusedHost, WM_KEYDOWN, vk, 0);
            SendMessageW(focusedHost, WM_KEYUP, vk, 0);
            result.closedAfterKey = WaitForWindowClosed(dialog, SelfTest::Scale(3000ms));
            if (! result.closedAfterKey && IsWindow(dialog) != FALSE)
            {
                if (accept)
                {
                    PostMessageW(dialog, WM_COMMAND, MAKEWPARAM(IDOK, 0), 0);
                }
                else
                {
                    static_cast<void>(DebugCancelPluginConfigurationDialog());
                }
                result.closedAfterFallback = WaitForWindowClosed(dialog, SelfTest::Scale(3000ms));
            }
        });

        Common::Settings::Settings baselineSettings = g_settings;
        Common::Settings::Settings workingSettings  = baselineSettings;
        const AppTheme theme = ResolveAppTheme(ThemeMode::Dark, std::format(L"plugin-config-{}-keyboard-selftest", accept ? L"accept" : L"cancel"));
        const HRESULT hr =
            EditPluginConfigurationDialog(mainWindow, PluginType::FileSystem, kBuiltinS3FileSystemId, entry->name, baselineSettings, workingSettings, theme);
        worker.join();

        state.Require(result.sawDialog, std::format(L"Plugin configuration dialog did not open during {}.", label));
        state.Require(result.ownedByMainWindow, std::format(L"Plugin configuration dialog should be owned by the main window during {}.", label));
        state.Require(result.capturedBaselineSnapshot, std::format(L"Failed to capture the plugin configuration baseline snapshot during {}.", label));
        state.Require(result.patternStats.has_value(), std::format(L"Failed to collect plugin configuration UI Automation pattern stats during {}.", label));
        if (result.patternStats.has_value())
        {
            state.Require(result.patternStats->editControlCount > 0u,
                          std::format(L"Plugin configuration dialog should expose visible DX edits before {}.", label));
            state.Require(result.patternStats->togglePatternCount > 0u,
                          std::format(L"Plugin configuration dialog should expose visible DX toggles before {}.", label));
            state.Require(result.patternStats->buttonControlCount > 0u && result.patternStats->invokePatternCount > 0u,
                          std::format(L"Plugin configuration dialog should expose visible DX command buttons before {}.", label));
        }
        state.Require(result.focusedFirstInput, std::format(L"Plugin configuration dialog did not focus the first visible DX input before {}.", label));
        state.Require(result.closedAfterKey,
                      std::format(L"Pressing {} from the focused plugin configuration DX input should close the dialog through {} routing.",
                                  accept ? L"Enter" : L"Escape",
                                  accept ? L"default-button" : L"cancel"));
        state.Require(hr == (accept ? S_OK : S_FALSE),
                      std::format(L"EditPluginConfigurationDialog returned unexpected HRESULT 0x{:08X} during {}.", static_cast<uint32_t>(hr), label));
        state.Require(GetPluginConfigurationDialogHandle() == nullptr || IsWindow(GetPluginConfigurationDialogHandle()) == FALSE,
                      std::format(L"Plugin configuration dialog should not remain open after {}.", label));
    };

    runPass(true, L"default-button Enter validation");
    if (! state.failure.empty())
    {
        return false;
    }

    runPass(false, L"cancel Escape validation");
    return state.failure.empty();
}

[[nodiscard]] bool TestPluginConfigurationDialogAccessKeysRouteExpectedActions(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;
    constexpr std::wstring_view kBuiltinS3FileSystemId = L"builtin/file-system-s3";

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    const FileSystemPluginManager::PluginEntry* entry = FindFileSystemPluginById(kBuiltinS3FileSystemId);
    state.Require(entry != nullptr, L"builtin/file-system-s3 plugin entry unavailable for access-key validation.");
    if (! entry)
    {
        return false;
    }

    struct WorkerResult
    {
        bool sawDialog                = false;
        bool capturedBaselineSnapshot = false;
        bool focusedFirstInput        = false;
        bool closedAfterMnemonic      = false;
        PluginConfigurationDialogDebugSnapshot baselineSnapshot{};
        PluginConfigurationDialogDebugSnapshot focusSnapshot{};
    } workerResult{};

    std::jthread worker([&](std::stop_token) noexcept
    {
        const HWND dialog      = WaitForWindow([] noexcept { return GetPluginConfigurationDialogHandle(); }, SelfTest::Scale(5000ms));
        workerResult.sawDialog = dialog != nullptr && IsWindow(dialog) != FALSE;
        if (! workerResult.sawDialog)
        {
            return;
        }

        const auto waitForSnapshot = [&](const auto& predicate, PluginConfigurationDialogDebugSnapshot& outSnapshot) noexcept
        {
            const auto deadline = std::chrono::steady_clock::now() + SelfTest::Scale(3000ms);
            while (std::chrono::steady_clock::now() < deadline)
            {
                PumpPendingMessages();
                outSnapshot = {};
                if (DebugGetPluginConfigurationDialogSnapshot(outSnapshot) && predicate(outSnapshot))
                {
                    return true;
                }
                std::this_thread::sleep_for(20ms);
            }

            outSnapshot = {};
            return DebugGetPluginConfigurationDialogSnapshot(outSnapshot) && predicate(outSnapshot);
        };

        workerResult.capturedBaselineSnapshot = waitForSnapshot(
            [](const PluginConfigurationDialogDebugSnapshot& value) noexcept
        {
            return value.usesDxUiCommandButtons && value.usesDxUiFormSurface && value.usesDxUiFormInputs && value.visibleLegacyCommandButtonCount == 0u &&
                   value.visibleLegacyFormControlCount == 0u && value.visibleDxCommandButtonHostCount == 2u && value.visibleDxFormInputHostCount > 0u;
        },
            workerResult.baselineSnapshot);
        if (! workerResult.capturedBaselineSnapshot)
        {
            static_cast<void>(DebugCancelPluginConfigurationDialog());
            return;
        }

        if (! DebugFocusPluginConfigurationDialogFirstInput())
        {
            static_cast<void>(DebugCancelPluginConfigurationDialog());
            return;
        }

        workerResult.focusedFirstInput = waitForSnapshot([](const PluginConfigurationDialogDebugSnapshot& value) noexcept {
            return value.focusKind == PluginConfigurationDialogDebugFocusKind::Edit && value.focusLabel == L"Default region";
        }, workerResult.focusSnapshot);
        if (! workerResult.focusedFirstInput)
        {
            static_cast<void>(DebugCancelPluginConfigurationDialog());
            return;
        }

        SendMessageW(dialog, WM_SYSCHAR, static_cast<WPARAM>(L'c'), 0);
        workerResult.closedAfterMnemonic = WaitForWindowClosed(dialog, SelfTest::Scale(3000ms));
        if (! workerResult.closedAfterMnemonic && IsWindow(dialog) != FALSE)
        {
            static_cast<void>(DebugCancelPluginConfigurationDialog());
            static_cast<void>(WaitForWindowClosed(dialog, SelfTest::Scale(3000ms)));
        }
    });

    Common::Settings::Settings baselineSettings = g_settings;
    Common::Settings::Settings workingSettings  = baselineSettings;
    const AppTheme theme                        = ResolveAppTheme(ThemeMode::Dark, L"plugin-config-access-keys-selftest");
    const HRESULT hr =
        EditPluginConfigurationDialog(mainWindow, PluginType::FileSystem, kBuiltinS3FileSystemId, entry->name, baselineSettings, workingSettings, theme);
    worker.join();

    state.Require(workerResult.sawDialog, L"Plugin configuration dialog did not open for access-key validation.");
    state.Require(workerResult.capturedBaselineSnapshot, L"Failed to capture the plugin configuration baseline snapshot before access-key validation.");
    state.Require(workerResult.focusedFirstInput, L"Plugin configuration dialog did not focus the first visible DX input before access-key validation.");
    state.Require(workerResult.closedAfterMnemonic, L"Plugin configuration dialog Alt+C should route through the visible DX Cancel action.");
    state.Require(hr == S_FALSE,
                  std::format(L"EditPluginConfigurationDialog returned unexpected HRESULT 0x{:08X} during access-key validation.", static_cast<uint32_t>(hr)));
    state.Require(GetPluginConfigurationDialogHandle() == nullptr || IsWindow(GetPluginConfigurationDialogHandle()) == FALSE,
                  L"Plugin configuration dialog should not remain open after access-key validation.");
    return state.failure.empty();
}

[[nodiscard]] bool TestPluginConfigurationDialogPointerClickTogglesVisibleDxToggle(HWND mainWindow, CaseState& state) noexcept
{
    using namespace std::chrono_literals;
    constexpr std::wstring_view kBuiltinS3FileSystemId = L"builtin/file-system-s3";

    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    const FileSystemPluginManager::PluginEntry* entry = FindFileSystemPluginById(kBuiltinS3FileSystemId);
    state.Require(entry != nullptr, L"builtin/file-system-s3 plugin entry unavailable for pointer-toggle validation.");
    if (! entry)
    {
        return false;
    }

    struct WorkerResult
    {
        bool sawDialog                = false;
        bool ownedByMainWindow        = false;
        bool capturedBaselineSnapshot = false;
        bool toggledOn                = false;
        bool toggledRestored          = false;
        bool cancelled                = false;
        PluginConfigurationDialogDebugSnapshot baselineSnapshot{};
        PluginConfigurationDialogDebugSnapshot finalSnapshot{};
        std::wstring toggleFailureDetails;
    } workerResult{};

    std::jthread worker([&](std::stop_token) noexcept
    {
        const HWND dialog      = WaitForWindow([] noexcept { return GetPluginConfigurationDialogHandle(); }, SelfTest::Scale(5000ms));
        workerResult.sawDialog = dialog != nullptr && IsWindow(dialog) != FALSE;
        if (! workerResult.sawDialog)
        {
            return;
        }

        workerResult.ownedByMainWindow = IsOwnedBy(dialog, mainWindow);

        const auto waitForSnapshot = [&](const auto& predicate, PluginConfigurationDialogDebugSnapshot& outSnapshot) noexcept
        {
            const auto deadline = std::chrono::steady_clock::now() + SelfTest::Scale(3000ms);
            while (std::chrono::steady_clock::now() < deadline)
            {
                PumpPendingMessages();
                outSnapshot = {};
                if (DebugGetPluginConfigurationDialogSnapshot(outSnapshot) && predicate(outSnapshot))
                {
                    return true;
                }
                std::this_thread::sleep_for(20ms);
            }

            outSnapshot = {};
            return DebugGetPluginConfigurationDialogSnapshot(outSnapshot) && predicate(outSnapshot);
        };

        workerResult.capturedBaselineSnapshot = waitForSnapshot(
            [](const PluginConfigurationDialogDebugSnapshot& value) noexcept
        {
            return value.usesDxUiCommandButtons && value.usesDxUiFormSurface && value.usesDxUiFormStatics && value.usesDxUiFormInputs &&
                   value.visibleLegacyCommandButtonCount == 0u && value.visibleLegacyFormControlCount == 0u && value.visibleDxCommandButtonHostCount == 2u &&
                   value.visibleDxFormInputHostCount > 0u && value.panelScrollPosY == 0;
        },
            workerResult.baselineSnapshot);
        if (! workerResult.capturedBaselineSnapshot)
        {
            static_cast<void>(DebugCancelPluginConfigurationDialog());
            return;
        }

        constexpr std::wstring_view kToggleLabel = L"Use HTTPS";
        const auto initialToggleState            = CollectVisibleDescendantTogglePatternState(dialog);
        if (! initialToggleState.has_value())
        {
            workerResult.toggleFailureDetails = L"no visible toggle descendant state was available for the Use HTTPS field.";
            static_cast<void>(DebugCancelPluginConfigurationDialog());
            return;
        }

        const ToggleState initialToggleValue = initialToggleState->toggleState;
        const ToggleState flippedToggleValue = (initialToggleValue == ToggleState_On) ? ToggleState_Off : ToggleState_On;
        HWND toggleHost                      = nullptr;
        RECT toggleRect{};
        if (! DebugGetPluginConfigurationDialogVisibleToggleHostAndClientRectByLabel(kToggleLabel, toggleHost, toggleRect) || ! toggleHost ||
            IsWindow(toggleHost) == FALSE)
        {
            static_cast<void>(DebugCancelPluginConfigurationDialog());
            return;
        }

        const int clickX        = toggleRect.left + ((toggleRect.right - toggleRect.left) / 2);
        const int clickY        = toggleRect.top + ((toggleRect.bottom - toggleRect.top) / 2);
        const LPARAM clickPoint = MAKELPARAM(clickX, clickY);

        const auto clickToggle = [&]() noexcept
        {
            SendMouseClickToResolvedPointWindow(toggleHost, clickPoint);
            PumpPendingMessages();
        };

        const auto requireToggleState = [&](const ToggleState expectedState) noexcept
        {
            const bool reached = waitForSnapshot(
                [&](const PluginConfigurationDialogDebugSnapshot& value) noexcept
            {
                const auto toggleState = CollectVisibleDescendantTogglePatternState(dialog);
                return toggleState.has_value() && toggleState->toggleState == expectedState &&
                       value.focusKind == PluginConfigurationDialogDebugFocusKind::Toggle && value.focusLabel == kToggleLabel &&
                       value.visibleDxCommandButtonHostCount == workerResult.baselineSnapshot.visibleDxCommandButtonHostCount &&
                       value.visibleDxFormStaticHostCount == workerResult.baselineSnapshot.visibleDxFormStaticHostCount &&
                       value.visibleDxFormInputHostCount == workerResult.baselineSnapshot.visibleDxFormInputHostCount &&
                       value.visibleLegacyCommandButtonCount == 0u && value.visibleLegacyFormControlCount == 0u && value.panelScrollPosY == 0;
            },
                workerResult.finalSnapshot);
            if (! reached)
            {
                const auto toggleState            = CollectVisibleDescendantTogglePatternState(dialog);
                workerResult.toggleFailureDetails = std::format(L"expectedState={} actualState={} focusKind={} focusLabel='{}' dxButtons={} dxStatics={} "
                                                                L"dxInputs={} legacyButtons={} legacyControls={} scrollY={} rect=({},{}-{}, {})",
                                                                static_cast<int>(expectedState),
                                                                toggleState.has_value() ? static_cast<int>(toggleState->toggleState) : -1,
                                                                static_cast<int>(workerResult.finalSnapshot.focusKind),
                                                                workerResult.finalSnapshot.focusLabel,
                                                                workerResult.finalSnapshot.visibleDxCommandButtonHostCount,
                                                                workerResult.finalSnapshot.visibleDxFormStaticHostCount,
                                                                workerResult.finalSnapshot.visibleDxFormInputHostCount,
                                                                workerResult.finalSnapshot.visibleLegacyCommandButtonCount,
                                                                workerResult.finalSnapshot.visibleLegacyFormControlCount,
                                                                workerResult.finalSnapshot.panelScrollPosY,
                                                                toggleRect.left,
                                                                toggleRect.top,
                                                                toggleRect.right,
                                                                toggleRect.bottom);
            }
            return reached;
        };

        clickToggle();
        workerResult.toggledOn = requireToggleState(flippedToggleValue);
        if (! workerResult.toggledOn)
        {
            static_cast<void>(DebugCancelPluginConfigurationDialog());
            return;
        }

        clickToggle();
        workerResult.toggledRestored = requireToggleState(initialToggleValue);
        workerResult.cancelled       = DebugCancelPluginConfigurationDialog();
    });

    Common::Settings::Settings baselineSettings = g_settings;
    Common::Settings::Settings workingSettings  = baselineSettings;
    const AppTheme theme                        = ResolveAppTheme(ThemeMode::Dark, L"plugin-config-pointer-toggle-selftest");
    const HRESULT hr =
        EditPluginConfigurationDialog(mainWindow, PluginType::FileSystem, kBuiltinS3FileSystemId, entry->name, baselineSettings, workingSettings, theme);
    worker.join();

    state.Require(workerResult.sawDialog, L"Plugin configuration dialog did not open for pointer-toggle validation.");
    state.Require(workerResult.ownedByMainWindow, L"Plugin configuration dialog should be owned by the main window during pointer-toggle validation.");
    state.Require(workerResult.capturedBaselineSnapshot, L"Failed to capture the plugin configuration baseline snapshot before pointer-toggle validation.");
    state.Require(workerResult.toggledOn,
                  workerResult.toggleFailureDetails.empty()
                      ? L"Plugin configuration dialog visible DX Use HTTPS toggle did not flip after the first real pointer click."
                      : std::format(L"Plugin configuration dialog visible DX Use HTTPS toggle did not flip after the first real pointer click. {}",
                                    workerResult.toggleFailureDetails));
    state.Require(workerResult.toggledRestored,
                  workerResult.toggleFailureDetails.empty()
                      ? L"Plugin configuration dialog visible DX Use HTTPS toggle did not restore after the second real pointer click."
                      : std::format(L"Plugin configuration dialog visible DX Use HTTPS toggle did not restore after the second real pointer click. {}",
                                    workerResult.toggleFailureDetails));
    state.Require(workerResult.cancelled, L"Plugin configuration dialog debug cancel command failed after pointer-toggle validation.");
    state.Require(
        hr == S_FALSE,
        std::format(L"EditPluginConfigurationDialog returned unexpected HRESULT 0x{:08X} during pointer-toggle validation.", static_cast<uint32_t>(hr)));
    return state.failure.empty();
}

[[nodiscard]] std::filesystem::path FindRepoRootForPluginManagerSourceGuard() noexcept
{
    std::error_code currentPathError;
    std::filesystem::path cursor = std::filesystem::current_path(currentPathError);
    if (currentPathError || cursor.empty())
    {
        return {};
    }

    for (int depth = 0; depth < 8; ++depth)
    {
        std::error_code ec;
        if (std::filesystem::exists(cursor / L"RedSalamander.sln", ec) && std::filesystem::exists(cursor / L"RedSalamander" / L"ViewerPluginManager.cpp", ec))
        {
            return cursor;
        }

        if (! cursor.has_parent_path())
        {
            break;
        }
        cursor = cursor.parent_path();
    }

    return {};
}

[[nodiscard]] bool ReadSourceFileUtf8(const std::filesystem::path& path, std::string& outSource) noexcept
{
    outSource.clear();
    std::ifstream input(path);
    if (! input.good())
    {
        return false;
    }

    outSource.assign(std::istreambuf_iterator<char>(input), std::istreambuf_iterator<char>());
    return true;
}

[[nodiscard]] std::string_view SourceBetween(std::string_view source, std::string_view beginMarker, std::string_view endMarker) noexcept
{
    const size_t begin = source.find(beginMarker);
    if (begin == std::string_view::npos)
    {
        return {};
    }

    const size_t end = source.find(endMarker, begin + beginMarker.size());
    if (end == std::string_view::npos || end <= begin)
    {
        return {};
    }

    return source.substr(begin, end - begin);
}

[[nodiscard]] bool TestViewerPluginManagerUnloadLifecycleSourceGuard(CaseState& state) noexcept
{
    const std::filesystem::path repoRoot = FindRepoRootForPluginManagerSourceGuard();
    state.Require(! repoRoot.empty(), L"Repository root unavailable for ViewerPluginManager source guard.");
    if (repoRoot.empty())
    {
        return false;
    }

    std::string source;
    const std::filesystem::path managerPath = repoRoot / L"RedSalamander" / L"ViewerPluginManager.cpp";
    state.Require(ReadSourceFileUtf8(managerPath, source), std::format(L"Failed to read {}.", managerPath.wstring()));
    if (source.empty())
    {
        return false;
    }

    const std::string_view disableBody =
        SourceBetween(source, "HRESULT ViewerPluginManager::DisablePlugin", "HRESULT ViewerPluginManager::EnablePlugin");
    state.Require(! disableBody.empty(), L"ViewerPluginManager::DisablePlugin body was not found.");
    state.Require(disableBody.find("Unload(") == std::string_view::npos,
                  L"ViewerPluginManager::DisablePlugin must not FreeLibrary a viewer DLL because external IViewer references may still be alive.");

    const std::string_view discoverBody = SourceBetween(source, "HRESULT ViewerPluginManager::Discover", "struct Candidate");
    state.Require(! discoverBody.empty(), L"ViewerPluginManager::Discover body was not found.");
    state.Require(source.find("void ViewerPluginManager::UnloadAll(") != std::string::npos,
                  L"ViewerPluginManager must centralize plugin unload through an explicit quiet-point helper.");
    state.Require(discoverBody.find("UnloadAll(ModuleUnloadMode::FreeLibrary);") != std::string_view::npos,
                  L"ViewerPluginManager::Discover must unload existing modules through the quiet-point helper before clearing plugin entries.");

    std::string appSource;
    const std::filesystem::path appPath = repoRoot / L"RedSalamander" / L"RedSalamander.cpp";
    state.Require(ReadSourceFileUtf8(appPath, appSource), std::format(L"Failed to read {}.", appPath.wstring()));
    const std::string_view refreshBody = SourceBetween(appSource, "void RefreshRunningPluginsFromSettings", "[[nodiscard]] std::vector<std::wstring_view>");
    state.Require(! refreshBody.empty(), L"RefreshRunningPluginsFromSettings body was not found.");
    const size_t closeAllViewersPos = refreshBody.find("g_folderWindow.CloseAllViewers();");
    const size_t releaseFileSystemsPos = refreshBody.find("g_folderWindow.ReleaseFileSystemPluginsForRefresh();");
    const size_t fileSystemRefreshPos = refreshBody.find("FileSystemPluginManager::GetInstance().Refresh");
    const size_t viewerRefreshPos     = refreshBody.find("ViewerPluginManager::GetInstance().Refresh");
    state.Require(closeAllViewersPos != std::string_view::npos && fileSystemRefreshPos != std::string_view::npos &&
                      viewerRefreshPos != std::string_view::npos && releaseFileSystemsPos != std::string_view::npos &&
                      closeAllViewersPos < releaseFileSystemsPos && releaseFileSystemsPos < fileSystemRefreshPos && closeAllViewersPos < viewerRefreshPos,
                  L"Runtime plugin refresh must close live viewers and release pane file-system references before rediscovering/unloading plugin DLLs.");

    return state.failure.empty();
}

[[nodiscard]] bool TestFileSystemPluginManagerUnloadLifecycleSourceGuard(CaseState& state) noexcept
{
    const std::filesystem::path repoRoot = FindRepoRootForPluginManagerSourceGuard();
    state.Require(! repoRoot.empty(), L"Repository root unavailable for FileSystemPluginManager source guard.");
    if (repoRoot.empty())
    {
        return false;
    }

    std::string source;
    const std::filesystem::path managerPath = repoRoot / L"RedSalamander" / L"FileSystemPluginManager.cpp";
    state.Require(ReadSourceFileUtf8(managerPath, source), std::format(L"Failed to read {}.", managerPath.wstring()));
    if (source.empty())
    {
        return false;
    }

    const std::string_view disableBody =
        SourceBetween(source, "HRESULT FileSystemPluginManager::DisablePlugin", "HRESULT FileSystemPluginManager::EnablePlugin");
    state.Require(! disableBody.empty(), L"FileSystemPluginManager::DisablePlugin body was not found.");
    state.Require(disableBody.find("Unload(") == std::string_view::npos,
                  L"FileSystemPluginManager::DisablePlugin must not FreeLibrary a file system DLL because external IFileSystem references may still be alive.");

    const std::string_view shutdownBody =
        SourceBetween(source, "void FileSystemPluginManager::Shutdown", "const std::vector<FileSystemPluginManager::PluginEntry>&");
    state.Require(! shutdownBody.empty(), L"FileSystemPluginManager::Shutdown body was not found.");
    state.Require(shutdownBody.find("UnloadAll(ModuleUnloadMode::ProcessShutdown);") != std::string_view::npos,
                  L"FileSystemPluginManager::Shutdown must unload modules through the process-shutdown quiet-point helper.");

    const std::string_view discoverBody = SourceBetween(source, "HRESULT FileSystemPluginManager::Discover", "    if (_exeDir.empty())");
    state.Require(! discoverBody.empty(), L"FileSystemPluginManager::Discover body was not found.");
    state.Require(source.find("void FileSystemPluginManager::UnloadAll(") != std::string::npos,
                  L"FileSystemPluginManager must centralize plugin unload through an explicit quiet-point helper.");
    state.Require(discoverBody.find("UnloadAll(ModuleUnloadMode::FreeLibrary);") != std::string_view::npos,
                  L"FileSystemPluginManager::Discover must unload existing modules through the quiet-point helper before clearing plugin entries.");

    const std::string_view unloadBody = SourceBetween(source, "void FileSystemPluginManager::Unload(", "HRESULT FileSystemPluginManager::ApplyConfigurationFromSettings");
    state.Require(! unloadBody.empty(), L"FileSystemPluginManager::Unload body was not found.");
    state.Require(unloadBody.find("RedSalamanderPluginShutdown") != std::string_view::npos,
                  L"FileSystemPluginManager::Unload must call optional plugin shutdown hooks before unloading.");
    state.Require(unloadBody.find("RedSalamanderPluginRetainModuleUntilProcessExit") != std::string_view::npos,
                  L"FileSystemPluginManager::Unload must honor process-shutdown module retention hooks.");

    return state.failure.empty();
}

} // namespace (tests)

void RunPluginConfigCommandsSelfTestCases(HWND mainWindow, const SelfTest::SelfTestOptions& options, SelfTest::SelfTestSuiteResult& suite) noexcept
{
    SelfTest::RunCase(
        options, suite, L"settings_file_system_plugin_roundtrip", [](CaseState& state) noexcept { return TestFileSystemPluginConfigurationRoundTrip(state); });
    SelfTest::RunCase(
        options, suite, L"settings_viewer_text_plugin_roundtrip", [](CaseState& state) noexcept { return TestViewerTextPluginConfigurationRoundTrip(state); });
    SelfTest::RunCase(
        options, suite, L"viewer_text_hex_byte_color_perf", [](CaseState& state) noexcept { return TestViewerTextHexByteColorPerfScenario(state); });
    SelfTest::RunCase(options, suite, L"viewer_text_diff_perf", [](CaseState& state) noexcept { return TestViewerTextDiffPerfScenario(state); });
    SelfTest::RunCase(options, suite, L"viewer_pe_close_roundtrip", [](CaseState& state) noexcept { return TestViewerPECloseRoundTrip(state); });
    SelfTest::RunCase(options, suite, L"viewer_imgraw_close_roundtrip", [](CaseState& state) noexcept { return TestViewerImgRawCloseRoundTrip(state); });
    SelfTest::RunCase(options, suite, L"viewer_web_close_roundtrip", [](CaseState& state) noexcept { return TestViewerWebCloseRoundTrip(state); });
    SelfTest::RunCase(options, suite, L"viewer_plugin_manager_unload_lifecycle_source_guard", [](CaseState& state) noexcept {
        return TestViewerPluginManagerUnloadLifecycleSourceGuard(state);
    });
    SelfTest::RunCase(options, suite, L"file_system_plugin_manager_unload_lifecycle_source_guard", [](CaseState& state) noexcept {
        return TestFileSystemPluginManagerUnloadLifecycleSourceGuard(state);
    });
    SelfTest::RunCase(options, suite, L"cmd_plugin_configuration_dialog_uses_dxui_command_buttons", [=](CaseState& state) noexcept {
        return TestPluginConfigurationDialogUsesDxUiFormSurface(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_plugin_configuration_dialog_uses_dxui_form_surface", [=](CaseState& state) noexcept {
        return TestPluginConfigurationDialogUsesDxUiFormSurface(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_plugin_configuration_dialog_long_run_scrolling_keeps_dx_surface_stable", [=](CaseState& state) noexcept {
        return TestPluginConfigurationDialogLongRunScrollingKeepsDxSurfaceStable(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_plugin_configuration_dialog_long_run_open_close_stays_stable", [=](CaseState& state) noexcept {
        return TestPluginConfigurationDialogLongRunOpenCloseStaysStable(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_plugin_configuration_dialog_theme_cycle_keeps_surface_legible", [=](CaseState& state) noexcept {
        return TestPluginConfigurationDialogThemeCycleKeepsSurfaceLegible(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_plugin_configuration_dialog_live_dx_interaction", [=](CaseState& state) noexcept {
        return TestPluginConfigurationDialogLiveDxInteraction(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_plugin_configuration_dialog_tab_traversal_live_dx_interaction", [=](CaseState& state) noexcept {
        return TestPluginConfigurationDialogTabTraversalLiveDxInteraction(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_plugin_configuration_dialog_enter_and_escape_route_default_cancel", [=](CaseState& state) noexcept {
        return TestPluginConfigurationDialogEnterAndEscapeRouteDefaultCancel(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_plugin_configuration_dialog_access_keys_route_expected_actions", [=](CaseState& state) noexcept {
        return TestPluginConfigurationDialogAccessKeysRouteExpectedActions(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"cmd_plugin_configuration_dialog_pointer_click_toggles_visible_dx_toggle", [=](CaseState& state) noexcept {
        return TestPluginConfigurationDialogPointerClickTogglesVisibleDxToggle(mainWindow, state);
    });
}

namespace
{
