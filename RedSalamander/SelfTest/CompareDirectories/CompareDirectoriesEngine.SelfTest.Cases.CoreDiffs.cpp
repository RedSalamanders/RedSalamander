SelfTest::RunCase(options,
                  suite,
                  L"windows_hello_cache",
                  [&](SelfTest::CaseState& state) noexcept
{
    const auto restoreVerifier = wil::scope_exit([&] noexcept { static_cast<void>(RedSalamander::Security::SetWindowsHelloTestVerifier(nullptr)); });
    g_windowsHelloVerifierCalls.store(0u, std::memory_order_relaxed);
    RedSalamander::Security::SetWindowsHelloTestVerifier(&TestWindowsHelloVerifier);

    std::wstring id = MakeGuidText();
    state.Require(! id.empty(), L"Failed to generate a connection profile id.");

    Common::Settings::ConnectionProfile profile;
    profile.id                  = id;
    profile.name                = std::format(L"SelfTest WindowsHello {}", id);
    profile.pluginId            = std::wstring(kBuiltinFtpFileSystemId);
    profile.host                = L"example.invalid";
    profile.initialPath         = L"/";
    profile.userName            = L"user";
    profile.authMode            = Common::Settings::ConnectionAuthMode::Password;
    profile.savePassword        = true;
    profile.requireWindowsHello = true;

    const bool hadConnectionsSettings = g_settings.connections.has_value();
    if (! hadConnectionsSettings)
    {
        g_settings.connections.emplace();
    }

    const bool previousBypassHello     = g_settings.connections->bypassWindowsHello;
    const uint32_t previousTimeoutMins = g_settings.connections->windowsHelloReauthTimeoutMinute;

    g_settings.connections->bypassWindowsHello              = false;
    g_settings.connections->windowsHelloReauthTimeoutMinute = 10;
    g_settings.connections->items.push_back(profile);

    const auto restoreSettings = wil::scope_exit([&] noexcept
    {
        RedSalamander::Connections::ClearSecretAccessAuthorization(id);

        if (! g_settings.connections)
        {
            return;
        }

        auto& items = g_settings.connections->items;
        items.erase(std::remove_if(items.begin(), items.end(), [&](const Common::Settings::ConnectionProfile& item) noexcept { return item.id == id; }),
                    items.end());

        g_settings.connections->bypassWindowsHello              = previousBypassHello;
        g_settings.connections->windowsHelloReauthTimeoutMinute = previousTimeoutMins;

        if (! hadConnectionsSettings)
        {
            g_settings.connections.reset();
        }
    });

    const std::wstring targetName = RedSalamander::Connections::BuildCredentialTargetName(profile.id, RedSalamander::Connections::SecretKind::Password);
    state.Require(! targetName.empty(), L"Failed to build WinCred target name.");

    const std::wstring password = L"pw";
    const HRESULT saveHr        = RedSalamander::Connections::SaveGenericCredential(targetName, profile.userName, password);
    state.Require(SUCCEEDED(saveHr), std::format(L"SaveGenericCredential failed. hr=0x{:08X}", static_cast<unsigned long>(saveHr)));
    const auto deleteCredential = wil::scope_exit([&] noexcept
    {
        const HRESULT delHr = RedSalamander::Connections::DeleteGenericCredential(targetName);
        if (FAILED(delHr) && delHr != HRESULT_FROM_WIN32(ERROR_NOT_FOUND))
        {
            Debug::Warning(L"SelfTest: DeleteGenericCredential failed target='{}' hr=0x{:08X}", targetName, static_cast<unsigned long>(delHr));
        }
    });

    RedSalamander::Connections::ClearSecretAccessAuthorization(profile.id);

    wil::com_ptr<IHostConnections> hostConnections;
    const HRESULT qiHr = GetHostServices()->QueryInterface(IID_PPV_ARGS(hostConnections.put()));
    state.Require(SUCCEEDED(qiHr) && hostConnections, std::format(L"Missing IHostConnections. hr=0x{:08X}", static_cast<unsigned long>(qiHr)));

    wil::unique_cotaskmem_string secret;
    HRESULT hr = hostConnections->GetConnectionSecret(profile.name.c_str(), HOST_CONNECTION_SECRET_PASSWORD, nullptr, secret.put());
    if (SkipIfHostConnectionUiUnavailable(state, hr, L"Windows Hello cache"))
    {
        return true;
    }
    state.Require(SUCCEEDED(hr), std::format(L"GetConnectionSecret failed. hr=0x{:08X}", static_cast<unsigned long>(hr)));
    state.Require(secret && std::wstring_view(secret.get()) == password, L"Unexpected secret value.");
    SecureClearAndFreeSecret(secret);
    state.Require(g_windowsHelloVerifierCalls.load(std::memory_order_relaxed) == 1u, L"Expected Windows Hello to be requested once.");

    wil::unique_cotaskmem_string secret2;
    hr = hostConnections->GetConnectionSecret(profile.name.c_str(), HOST_CONNECTION_SECRET_PASSWORD, nullptr, secret2.put());
    if (SkipIfHostConnectionUiUnavailable(state, hr, L"Windows Hello cache second read"))
    {
        return true;
    }
    state.Require(SUCCEEDED(hr), std::format(L"GetConnectionSecret (second call) failed. hr=0x{:08X}", static_cast<unsigned long>(hr)));
    SecureClearAndFreeSecret(secret2);
    state.Require(g_windowsHelloVerifierCalls.load(std::memory_order_relaxed) == 1u, L"Expected Windows Hello to be cached (no second prompt).");

#ifdef ENABLE_TESTS
    // Simulate an expired authorization timestamp to ensure long-running operations (compare/copy) won't re-prompt.
    RedSalamander::Connections::SetSecretAccessAuthorizationTickForTesting(
        profile.id, RedSalamander::Connections::SecretKind::Password, RedSalamander::Connections::SecretAccessPurpose::Interactive, 0u);

    wil::unique_cotaskmem_string secretExpired;
    hr = hostConnections->GetConnectionSecret(profile.name.c_str(), HOST_CONNECTION_SECRET_PASSWORD, nullptr, secretExpired.put());
    if (SkipIfHostConnectionUiUnavailable(state, hr, L"Windows Hello cache expired-auth read"))
    {
        return true;
    }
    state.Require(SUCCEEDED(hr), std::format(L"GetConnectionSecret (expired auth) failed. hr=0x{:08X}", static_cast<unsigned long>(hr)));
    SecureClearAndFreeSecret(secretExpired);
    state.Require(g_windowsHelloVerifierCalls.load(std::memory_order_relaxed) == 1u, L"Expected Windows Hello not to re-prompt after session auth.");
#endif

    RedSalamander::Connections::ClearSecretAccessAuthorization(profile.id);
    RedSalamander::Connections::NoteSecretAccessAuthorized(profile.id, RedSalamander::Connections::SecretKind::Password);
    g_windowsHelloVerifierCalls.store(0u, std::memory_order_relaxed);

    wil::unique_cotaskmem_string secret3;
    hr = hostConnections->GetConnectionSecret(profile.name.c_str(), HOST_CONNECTION_SECRET_PASSWORD, nullptr, secret3.put());
    if (SkipIfHostConnectionUiUnavailable(state, hr, L"Windows Hello cache manual-auth read"))
    {
        return true;
    }
    state.Require(SUCCEEDED(hr), std::format(L"GetConnectionSecret (manual auth) failed. hr=0x{:08X}", static_cast<unsigned long>(hr)));
    SecureClearAndFreeSecret(secret3);
    state.Require(g_windowsHelloVerifierCalls.load(std::memory_order_relaxed) == 0u, L"Manual secret entry should suppress Windows Hello prompts.");

    return state.failure.empty();
});

SelfTest::RunCase(options,
                  suite,
                  L"oauth_refresh_token_storage",
                  [&](SelfTest::CaseState& state) noexcept
{
    const auto restoreVerifier = wil::scope_exit([&] noexcept { static_cast<void>(RedSalamander::Security::SetWindowsHelloTestVerifier(nullptr)); });
    g_windowsHelloVerifierCalls.store(0u, std::memory_order_relaxed);
    RedSalamander::Security::SetWindowsHelloTestVerifier(&TestWindowsHelloVerifier);

    std::wstring id = MakeGuidText();
    state.Require(! id.empty(), L"Failed to generate a connection profile id.");

    Common::Settings::ConnectionProfile profile;
    profile.id                  = id;
    profile.name                = std::format(L"SelfTest GoogleDrive {}", id);
    profile.pluginId            = L"builtin/file-system-gdrive";
    profile.host                = L"";
    profile.initialPath         = L"/";
    profile.userName            = L"user@example.invalid";
    profile.authMode            = Common::Settings::ConnectionAuthMode::OAuth2Pkce;
    profile.savePassword        = true;
    profile.requireWindowsHello = true;

    const bool hadConnectionsSettings = g_settings.connections.has_value();
    if (! hadConnectionsSettings)
    {
        g_settings.connections.emplace();
    }

    const bool previousBypassHello     = g_settings.connections->bypassWindowsHello;
    const uint32_t previousTimeoutMins = g_settings.connections->windowsHelloReauthTimeoutMinute;

    g_settings.connections->bypassWindowsHello              = false;
    g_settings.connections->windowsHelloReauthTimeoutMinute = 10;
    g_settings.connections->items.push_back(profile);

    const auto restoreSettings = wil::scope_exit([&] noexcept
    {
        RedSalamander::Connections::ClearSecretAccessAuthorization(id);

        wil::com_ptr<IHostConnections> cleanupHostConnections;
        if (SUCCEEDED(GetHostServices()->QueryInterface(IID_PPV_ARGS(cleanupHostConnections.put()))) && cleanupHostConnections)
        {
            static_cast<void>(cleanupHostConnections->DeleteConnectionSecret(profile.name.c_str(), HOST_CONNECTION_SECRET_OAUTH_REFRESH_TOKEN, TRUE));
            static_cast<void>(cleanupHostConnections->ClearCachedConnectionSecret(profile.name.c_str(), HOST_CONNECTION_SECRET_OAUTH_REFRESH_TOKEN));
        }

        if (! g_settings.connections)
        {
            return;
        }

        auto& items = g_settings.connections->items;
        items.erase(std::remove_if(items.begin(), items.end(), [&](const Common::Settings::ConnectionProfile& item) noexcept { return item.id == id; }),
                    items.end());

        g_settings.connections->bypassWindowsHello              = previousBypassHello;
        g_settings.connections->windowsHelloReauthTimeoutMinute = previousTimeoutMins;

        if (! hadConnectionsSettings)
        {
            g_settings.connections.reset();
        }
    });

    wil::com_ptr<IHostConnections> hostConnections;
    const HRESULT qiHr = GetHostServices()->QueryInterface(IID_PPV_ARGS(hostConnections.put()));
    state.Require(SUCCEEDED(qiHr) && hostConnections, std::format(L"Missing IHostConnections. hr=0x{:08X}", static_cast<unsigned long>(qiHr)));

    const std::wstring refreshToken = L"refresh-token-selftest";
    HRESULT hr = hostConnections->SetConnectionSecret(profile.name.c_str(), HOST_CONNECTION_SECRET_OAUTH_REFRESH_TOKEN, refreshToken.c_str(), TRUE);
    if (SkipIfHostConnectionUiUnavailable(state, hr, L"OAuth refresh token persisted store"))
    {
        return true;
    }
    state.Require(SUCCEEDED(hr), std::format(L"SetConnectionSecret failed. hr=0x{:08X}", static_cast<unsigned long>(hr)));

    hr = hostConnections->ClearCachedConnectionSecret(profile.name.c_str(), HOST_CONNECTION_SECRET_OAUTH_REFRESH_TOKEN);
    if (SkipIfHostConnectionUiUnavailable(state, hr, L"OAuth refresh token persisted cache clear"))
    {
        return true;
    }
    state.Require(SUCCEEDED(hr), std::format(L"ClearCachedConnectionSecret failed. hr=0x{:08X}", static_cast<unsigned long>(hr)));

    wil::unique_cotaskmem_string secret;
    hr = hostConnections->GetConnectionSecret(profile.name.c_str(), HOST_CONNECTION_SECRET_OAUTH_REFRESH_TOKEN, nullptr, secret.put());
    if (SkipIfHostConnectionUiUnavailable(state, hr, L"OAuth refresh token persisted read"))
    {
        return true;
    }
    state.Require(SUCCEEDED(hr), std::format(L"GetConnectionSecret failed. hr=0x{:08X}", static_cast<unsigned long>(hr)));
    state.Require(secret && std::wstring_view(secret.get()) == refreshToken, L"Unexpected OAuth refresh token value.");
    SecureClearAndFreeSecret(secret);
    state.Require(g_windowsHelloVerifierCalls.load(std::memory_order_relaxed) == 1u, L"Expected Windows Hello to guard persisted refresh tokens.");

    hr = hostConnections->DeleteConnectionSecret(profile.name.c_str(), HOST_CONNECTION_SECRET_OAUTH_REFRESH_TOKEN, TRUE);
    if (SkipIfHostConnectionUiUnavailable(state, hr, L"OAuth refresh token persisted delete"))
    {
        return true;
    }
    state.Require(SUCCEEDED(hr), std::format(L"DeleteConnectionSecret failed. hr=0x{:08X}", static_cast<unsigned long>(hr)));

    hr = hostConnections->ClearCachedConnectionSecret(profile.name.c_str(), HOST_CONNECTION_SECRET_OAUTH_REFRESH_TOKEN);
    if (SkipIfHostConnectionUiUnavailable(state, hr, L"OAuth refresh token post-delete cache clear"))
    {
        return true;
    }
    state.Require(SUCCEEDED(hr), std::format(L"ClearCachedConnectionSecret (post-delete) failed. hr=0x{:08X}", static_cast<unsigned long>(hr)));

    wil::unique_cotaskmem_string missingSecret;
    hr = hostConnections->GetConnectionSecret(profile.name.c_str(), HOST_CONNECTION_SECRET_OAUTH_REFRESH_TOKEN, nullptr, missingSecret.put());
    if (SkipIfHostConnectionUiUnavailable(state, hr, L"OAuth refresh token post-delete read"))
    {
        return true;
    }
    state.Require(hr == HRESULT_FROM_WIN32(ERROR_NOT_FOUND),
                  std::format(L"Expected ERROR_NOT_FOUND after deleting saved refresh token. hr=0x{:08X}", static_cast<unsigned long>(hr)));

    const std::wstring sessionRefreshToken = L"session-refresh-token";
    hr = hostConnections->SetConnectionSecret(profile.name.c_str(), HOST_CONNECTION_SECRET_OAUTH_REFRESH_TOKEN, sessionRefreshToken.c_str(), FALSE);
    if (SkipIfHostConnectionUiUnavailable(state, hr, L"OAuth refresh token session store"))
    {
        return true;
    }
    state.Require(SUCCEEDED(hr), std::format(L"SetConnectionSecret (session only) failed. hr=0x{:08X}", static_cast<unsigned long>(hr)));

    wil::unique_cotaskmem_string sessionSecret;
    hr = hostConnections->GetConnectionSecret(profile.name.c_str(), HOST_CONNECTION_SECRET_OAUTH_REFRESH_TOKEN, nullptr, sessionSecret.put());
    if (SkipIfHostConnectionUiUnavailable(state, hr, L"OAuth refresh token session read"))
    {
        return true;
    }
    state.Require(SUCCEEDED(hr), std::format(L"GetConnectionSecret (session only) failed. hr=0x{:08X}", static_cast<unsigned long>(hr)));
    state.Require(sessionSecret && std::wstring_view(sessionSecret.get()) == sessionRefreshToken, L"Unexpected session OAuth refresh token value.");
    SecureClearAndFreeSecret(sessionSecret);

    hr = hostConnections->ClearCachedConnectionSecret(profile.name.c_str(), HOST_CONNECTION_SECRET_OAUTH_REFRESH_TOKEN);
    if (SkipIfHostConnectionUiUnavailable(state, hr, L"OAuth refresh token session cache clear"))
    {
        return true;
    }
    state.Require(SUCCEEDED(hr), std::format(L"ClearCachedConnectionSecret (session only) failed. hr=0x{:08X}", static_cast<unsigned long>(hr)));

    wil::unique_cotaskmem_string sessionMissing;
    hr = hostConnections->GetConnectionSecret(profile.name.c_str(), HOST_CONNECTION_SECRET_OAUTH_REFRESH_TOKEN, nullptr, sessionMissing.put());
    if (SkipIfHostConnectionUiUnavailable(state, hr, L"OAuth refresh token session missing read"))
    {
        return true;
    }
    state.Require(hr == HRESULT_FROM_WIN32(ERROR_NOT_FOUND),
                  std::format(L"Expected ERROR_NOT_FOUND after clearing session-only refresh token. hr=0x{:08X}", static_cast<unsigned long>(hr)));

#ifdef ENABLE_TESTS
    const std::wstring durableOldToken = L"durable-old-refresh-token";
    hr = hostConnections->SetConnectionSecret(profile.name.c_str(), HOST_CONNECTION_SECRET_OAUTH_REFRESH_TOKEN, durableOldToken.c_str(), TRUE);
    state.Require(SUCCEEDED(hr), std::format(L"Failed to seed durable token for persistence-failure proof. hr=0x{:08X}", static_cast<unsigned long>(hr)));

    const std::wstring rotatedSessionToken = L"rotated-session-token";
    RedSalamander::Connections::SetCredentialPersistenceFaultForTesting(RedSalamander::Connections::CredentialPersistenceFault::SaveOnce);
    hr = hostConnections->SetConnectionSecret(profile.name.c_str(), HOST_CONNECTION_SECRET_OAUTH_REFRESH_TOKEN, rotatedSessionToken.c_str(), TRUE);
    state.Require(hr == HRESULT_FROM_WIN32(ERROR_ACCESS_DENIED),
                  std::format(L"Injected credential save failure expected ACCESS_DENIED, got 0x{:08X}", static_cast<unsigned long>(hr)));

    wil::unique_cotaskmem_string rotatedSecret;
    hr = hostConnections->GetConnectionSecret(profile.name.c_str(), HOST_CONNECTION_SECRET_OAUTH_REFRESH_TOKEN, nullptr, rotatedSecret.put());
    state.Require(SUCCEEDED(hr) && rotatedSecret && std::wstring_view(rotatedSecret.get()) == rotatedSessionToken,
                  L"Credential save failure exposed the stale durable token instead of the rotated session token.");
    SecureClearAndFreeSecret(rotatedSecret);

    RedSalamander::Connections::SetCredentialPersistenceFaultForTesting(RedSalamander::Connections::CredentialPersistenceFault::DeleteOnce);
    hr = hostConnections->DeleteConnectionSecret(profile.name.c_str(), HOST_CONNECTION_SECRET_OAUTH_REFRESH_TOKEN, TRUE);
    state.Require(hr == HRESULT_FROM_WIN32(ERROR_ACCESS_DENIED),
                  std::format(L"Injected credential delete failure expected ACCESS_DENIED, got 0x{:08X}", static_cast<unsigned long>(hr)));

    wil::unique_cotaskmem_string deletedSessionSecret;
    hr = hostConnections->GetConnectionSecret(profile.name.c_str(), HOST_CONNECTION_SECRET_OAUTH_REFRESH_TOKEN, nullptr, deletedSessionSecret.put());
    state.Require(hr == HRESULT_FROM_WIN32(ERROR_NOT_FOUND) && ! deletedSessionSecret,
                  L"Credential delete failure re-exposed the stale durable token after the session tombstone was installed.");
#endif

    return state.failure.empty();
});

SelfTest::RunCase(options,
                  suite,
                  L"oauth_authmode_roundtrip",
                  [&](SelfTest::CaseState& state) noexcept
{
    const std::wstring id = MakeGuidText();
    state.Require(! id.empty(), L"Failed to generate a connection profile id.");

    Common::Settings::Settings settings;
    settings.connections.emplace();

    Common::Settings::ConnectionProfile profile;
    profile.id                  = id;
    profile.name                = std::format(L"Settings OAuth {}", id);
    profile.pluginId            = L"builtin/file-system-gdrive";
    profile.host                = L"";
    profile.initialPath         = L"/";
    profile.userName            = L"user@example.invalid";
    profile.authMode            = Common::Settings::ConnectionAuthMode::OAuth2Pkce;
    profile.savePassword        = true;
    profile.requireWindowsHello = true;
    settings.connections->items.push_back(profile);

    const std::wstring appId                 = std::format(L"RedSalamanderSelfTestOAuth{}", id);
    const std::filesystem::path settingsPath = Common::Settings::GetSettingsPath(appId);
    const std::filesystem::path schemaPath   = Common::Settings::GetSettingsSchemaPath(appId);
    const auto cleanupSettings               = wil::scope_exit([&] noexcept
    {
        std::error_code ec;
        std::filesystem::remove(settingsPath, ec);
        std::filesystem::remove(schemaPath, ec);
    });

    const HRESULT saveHr = Common::Settings::SaveSettings(appId, settings);
    if (! state.Require(saveHr == S_OK, std::format(L"SaveSettings failed. hr=0x{:08X}", static_cast<unsigned long>(saveHr))))
    {
        return false;
    }

    Common::Settings::Settings loaded;
    const HRESULT loadHr = Common::Settings::LoadSettings(appId, loaded);
    if (! state.Require(loadHr == S_OK, std::format(L"LoadSettings failed. hr=0x{:08X}", static_cast<unsigned long>(loadHr))))
    {
        return false;
    }
    if (! state.Require(loaded.connections.has_value(), L"Expected connections settings after reload."))
    {
        return false;
    }
    if (! state.Require(! loaded.connections->items.empty(), L"Expected one connection profile after reload."))
    {
        return false;
    }
    state.Require(loaded.connections->items.front().authMode == Common::Settings::ConnectionAuthMode::OAuth2Pkce,
                  L"Expected OAuth2 PKCE auth mode after settings round-trip.");

    return state.failure.empty();
});

SelfTest::RunCase(options,
                  suite,
                  L"google_drive_plugin_contract",
                  [&](SelfTest::CaseState& state) noexcept
{
    CreatedFileSystemInstance created{};
    const HRESULT createHr = TryCreateFileSystemInstance(kBuiltinGoogleDriveFileSystemId, {}, created);
    state.Require(SUCCEEDED(createHr) && created.fileSystem,
                  std::format(L"Google Drive plugin: failed to create filesystem instance. hr=0x{:08X}", static_cast<unsigned long>(createHr)));
    if (FAILED(createHr) || ! created.fileSystem)
    {
        return false;
    }

    wil::com_ptr<IInformations> info;
    state.Require(CreateInformations(created.fileSystem, info), L"Google Drive plugin: missing IInformations.");
    if (! info)
    {
        return false;
    }

    const PluginMetaData* meta = nullptr;
    HRESULT hr                 = info->GetMetaData(&meta);
    state.Require(SUCCEEDED(hr) && meta, std::format(L"Google Drive plugin: GetMetaData failed. hr=0x{:08X}", static_cast<unsigned long>(hr)));
    if (FAILED(hr) || ! meta)
    {
        return false;
    }

    state.Require(meta->id && std::wstring_view(meta->id) == kBuiltinGoogleDriveFileSystemId, L"Google Drive plugin: unexpected metadata id.");
    state.Require(meta->shortId && std::wstring_view(meta->shortId) == L"gdrive", L"Google Drive plugin: unexpected shortId.");
    state.Require(meta->name && std::wstring_view(meta->name) == L"Google Drive", L"Google Drive plugin: unexpected display name.");

    const char* schema = nullptr;
    hr                 = info->GetConfigurationSchema(&schema);
    state.Require(SUCCEEDED(hr) && schema, std::format(L"Google Drive plugin: GetConfigurationSchema failed. hr=0x{:08X}", static_cast<unsigned long>(hr)));
    if (FAILED(hr) || ! schema)
    {
        return false;
    }

    const std::string_view schemaView(schema);
    state.Require(schemaView.find("\"defaultClientId\"") != std::string_view::npos, L"Google Drive plugin: schema missing defaultClientId.");
    state.Require(schemaView.find("\"pageSize\"") != std::string_view::npos, L"Google Drive plugin: schema missing pageSize.");

    BOOL somethingToSave = TRUE;
    hr                   = info->SomethingToSave(&somethingToSave);
    state.Require(SUCCEEDED(hr) && somethingToSave == FALSE,
                  std::format(L"Google Drive plugin: SomethingToSave before configuration failed. hr=0x{:08X}", static_cast<unsigned long>(hr)));

    hr = info->SetConfiguration(
        "{\"defaultClientId\":\"selftest-client-id.apps.googleusercontent.com\",\"connectTimeoutMs\":1234,\"requestTimeoutMs\":5678,\"pageSize\":321}");
    state.Require(SUCCEEDED(hr), std::format(L"Google Drive plugin: SetConfiguration failed. hr=0x{:08X}", static_cast<unsigned long>(hr)));
    if (FAILED(hr))
    {
        return false;
    }

    hr = info->SomethingToSave(&somethingToSave);
    state.Require(SUCCEEDED(hr) && somethingToSave == TRUE,
                  std::format(L"Google Drive plugin: SomethingToSave after configuration failed. hr=0x{:08X}", static_cast<unsigned long>(hr)));

    const char* configuration = nullptr;
    hr                        = info->GetConfiguration(&configuration);
    state.Require(SUCCEEDED(hr) && configuration, std::format(L"Google Drive plugin: GetConfiguration failed. hr=0x{:08X}", static_cast<unsigned long>(hr)));
    if (FAILED(hr) || ! configuration)
    {
        return false;
    }

    const std::string_view configurationView(configuration);
    state.Require(configurationView.find("selftest-client-id.apps.googleusercontent.com") != std::string_view::npos,
                  L"Google Drive plugin: configuration missing defaultClientId.");
    state.Require(configurationView.find("\"pageSize\":321") != std::string_view::npos, L"Google Drive plugin: configuration missing pageSize.");

    const char* capabilities = nullptr;
    hr                       = created.fileSystem->GetCapabilities(&capabilities);
    state.Require(SUCCEEDED(hr) && capabilities, std::format(L"Google Drive plugin: GetCapabilities failed. hr=0x{:08X}", static_cast<unsigned long>(hr)));
    if (FAILED(hr) || ! capabilities)
    {
        return false;
    }

    const std::string_view capabilitiesView(capabilities);
    state.Require(capabilitiesView.find("\"copy\": false") != std::string_view::npos, L"Google Drive plugin: capabilities should report copy=false.");
    state.Require(capabilitiesView.find("\"delete\": false") != std::string_view::npos, L"Google Drive plugin: capabilities should report delete=false.");

    wil::com_ptr<INavigationMenu> navigationMenu;
    hr = created.fileSystem->QueryInterface(IID_PPV_ARGS(navigationMenu.put()));
    state.Require(SUCCEEDED(hr) && navigationMenu, std::format(L"Google Drive plugin: missing INavigationMenu. hr=0x{:08X}", static_cast<unsigned long>(hr)));
    if (FAILED(hr) || ! navigationMenu)
    {
        return false;
    }

    const NavigationMenuItem* menuItems = nullptr;
    unsigned int menuCount              = 0;
    hr                                  = navigationMenu->GetMenuItems(&menuItems, &menuCount);
    state.Require(SUCCEEDED(hr) && menuItems && menuCount >= 4u,
                  std::format(L"Google Drive plugin: GetMenuItems failed. hr=0x{:08X}", static_cast<unsigned long>(hr)));
    if (FAILED(hr) || ! menuItems || menuCount < 4u)
    {
        return false;
    }

    state.Require((menuItems[0].flags & NAV_MENU_ITEM_FLAG_HEADER) != 0, L"Google Drive plugin: first menu item should be a header.");
    state.Require((menuItems[1].flags & NAV_MENU_ITEM_FLAG_SEPARATOR) != 0, L"Google Drive plugin: second menu item should be a separator.");
    state.Require(menuItems[2].commandId == 1u && menuItems[2].path == nullptr,
                  L"Google Drive plugin: third menu item should be the connection-manager command.");
    state.Require(menuItems[3].label && std::wstring_view(menuItems[3].label) == L"/", L"Google Drive plugin: root menu item label mismatch.");
    state.Require(menuItems[3].path && std::wstring_view(menuItems[3].path) == L"/", L"Google Drive plugin: root menu item path mismatch.");

    wil::com_ptr<IDriveInfo> driveInfoService;
    hr = created.fileSystem->QueryInterface(IID_PPV_ARGS(driveInfoService.put()));
    state.Require(SUCCEEDED(hr) && driveInfoService, std::format(L"Google Drive plugin: missing IDriveInfo. hr=0x{:08X}", static_cast<unsigned long>(hr)));
    if (FAILED(hr) || ! driveInfoService)
    {
        return false;
    }

    DriveInfo drive{};
    hr = driveInfoService->GetDriveInfo(L"/", &drive);
    state.Require(SUCCEEDED(hr), std::format(L"Google Drive plugin: GetDriveInfo('/') failed. hr=0x{:08X}", static_cast<unsigned long>(hr)));
    if (FAILED(hr))
    {
        return false;
    }

    state.Require((drive.flags & DRIVE_INFO_FLAG_HAS_DISPLAY_NAME) != 0, L"Google Drive plugin: root drive info should expose a display name.");
    state.Require(drive.displayName && std::wstring_view(drive.displayName) == L"gdrive:/", L"Google Drive plugin: root drive display name mismatch.");
    state.Require((drive.flags & DRIVE_INFO_FLAG_HAS_VOLUME_LABEL) != 0, L"Google Drive plugin: root drive info should expose a volume label.");
    state.Require(drive.volumeLabel && std::wstring_view(drive.volumeLabel) == L"Google Drive", L"Google Drive plugin: root drive volume label mismatch.");

    const NavigationMenuItem* driveMenuItems = reinterpret_cast<const NavigationMenuItem*>(static_cast<uintptr_t>(1));
    unsigned int driveMenuCount              = 99;
    hr                                       = driveInfoService->GetDriveMenuItems(L"/", &driveMenuItems, &driveMenuCount);
    state.Require(SUCCEEDED(hr) && driveMenuItems == nullptr && driveMenuCount == 0u,
                  std::format(L"Google Drive plugin: GetDriveMenuItems failed. hr=0x{:08X}", static_cast<unsigned long>(hr)));

    wil::com_ptr<IFileSystemIO> io;
    const HRESULT ioHr = created.fileSystem->QueryInterface(IID_PPV_ARGS(io.put()));
    state.Require(ioHr == E_NOINTERFACE && ! io, L"Google Drive plugin: IFileSystemIO should be absent in the current read-only milestone.");

    wil::com_ptr<IFilesInformation> filesInformation;
    hr = created.fileSystem->ReadDirectoryInfo(L"/", filesInformation.put());
    state.Require(SUCCEEDED(hr) && filesInformation,
                  std::format(L"Google Drive plugin: ReadDirectoryInfo('/') failed. hr=0x{:08X}", static_cast<unsigned long>(hr)));
    if (FAILED(hr) || ! filesInformation)
    {
        return false;
    }

    unsigned long entryCount = 99;
    hr                       = filesInformation->GetCount(&entryCount);
    state.Require(SUCCEEDED(hr) && entryCount == 0u,
                  std::format(L"Google Drive plugin: GetCount on root listing failed. hr=0x{:08X}", static_cast<unsigned long>(hr)));

    return state.failure.empty();
});

SelfTest::RunCase(options,
                  suite,
                  L"google_drive_navigation_menu_callback_clear_drains",
                  [&](SelfTest::CaseState& state) noexcept
{
    wil::com_ptr<BlockingConnectionManagerHost> host;
    host.attach(new (std::nothrow) BlockingConnectionManagerHost());
    state.Require(host != nullptr, L"Google Drive navigation callback drain: failed to create fake host.");
    if (! host)
    {
        return false;
    }

    CreatedFileSystemInstance created{};
    const HRESULT createHr = TryCreateFileSystemInstanceWithHost(kBuiltinGoogleDriveFileSystemId, static_cast<IHost*>(host.get()), {}, created);
    state.Require(
        SUCCEEDED(createHr) && created.fileSystem,
        std::format(L"Google Drive navigation callback drain: failed to create filesystem instance. hr=0x{:08X}", static_cast<unsigned long>(createHr)));
    if (FAILED(createHr) || ! created.fileSystem)
    {
        return false;
    }

    wil::com_ptr<INavigationMenu> navigationMenu;
    const HRESULT qiHr = created.fileSystem->QueryInterface(IID_PPV_ARGS(navigationMenu.put()));
    state.Require(SUCCEEDED(qiHr) && navigationMenu,
                  std::format(L"Google Drive navigation callback drain: missing INavigationMenu. hr=0x{:08X}", static_cast<unsigned long>(qiHr)));
    if (FAILED(qiHr) || ! navigationMenu)
    {
        return false;
    }

    NavigationMenuCallbackProbe callbackProbe;
    void* const callbackCookie = &callbackProbe;
    HRESULT hr                 = navigationMenu->SetCallback(&callbackProbe, callbackCookie);
    state.Require(SUCCEEDED(hr),
                  std::format(L"Google Drive navigation callback drain: SetCallback registration failed. hr=0x{:08X}", static_cast<unsigned long>(hr)));
    if (FAILED(hr))
    {
        return false;
    }

    host->PrepareShowConnectionManagerResult(L"Navigation Callback SelfTest");
    HRESULT normalCommandHr = E_PENDING;
    std::jthread normalCommandThread([&]() noexcept { normalCommandHr = navigationMenu->ExecuteMenuCommand(1u); });
    auto releaseNormalCommand = wil::scope_exit([&]() noexcept
    {
        host->ReleaseShowConnectionManager();
        if (normalCommandThread.joinable())
        {
            normalCommandThread.join();
        }
    });

    state.Require(host->WaitForShowConnectionManagerCall(std::chrono::seconds(5)),
                  L"Google Drive navigation callback drain: ExecuteMenuCommand did not reach ShowConnectionManager.");
    if (! state.failure.empty())
    {
        return false;
    }

    host->ReleaseShowConnectionManager();
    normalCommandThread.join();
    state.Require(
        SUCCEEDED(normalCommandHr),
        std::format(L"Google Drive navigation callback drain: normal ExecuteMenuCommand failed. hr=0x{:08X}", static_cast<unsigned long>(normalCommandHr)));
    state.Require(callbackProbe.GetCallCount() == 1u, L"Google Drive navigation callback drain: expected one callback before clear.");
    state.Require(callbackProbe.GetLastPath() == L"/@conn:Navigation Callback SelfTest/",
                  L"Google Drive navigation callback drain: callback path mismatch before clear.");
    state.Require(callbackProbe.GetLastCookie() == callbackCookie, L"Google Drive navigation callback drain: callback cookie mismatch before clear.");
    if (! state.failure.empty())
    {
        return false;
    }

    callbackProbe.Reset();
    hr = navigationMenu->SetCallback(&callbackProbe, callbackCookie);
    state.Require(SUCCEEDED(hr),
                  std::format(L"Google Drive navigation callback drain: SetCallback re-registration failed. hr=0x{:08X}", static_cast<unsigned long>(hr)));
    if (FAILED(hr))
    {
        return false;
    }

    host->PrepareShowConnectionManagerResult(L"Navigation Callback Cleared");
    HRESULT staleCommandHr = E_PENDING;
    std::jthread staleCommandThread([&]() noexcept { staleCommandHr = navigationMenu->ExecuteMenuCommand(1u); });
    auto releaseStaleCommand = wil::scope_exit([&]() noexcept
    {
        host->ReleaseShowConnectionManager();
        if (staleCommandThread.joinable())
        {
            staleCommandThread.join();
        }
    });

    state.Require(host->WaitForShowConnectionManagerCall(std::chrono::seconds(5)),
                  L"Google Drive navigation callback drain: stale ExecuteMenuCommand did not reach ShowConnectionManager.");
    if (! state.failure.empty())
    {
        return false;
    }

    std::mutex clearMutex;
    std::condition_variable clearCv;
    bool clearCompleted = false;
    HRESULT clearHr     = E_PENDING;
    std::jthread clearThread([&]() noexcept
    {
        const HRESULT localHr = navigationMenu->SetCallback(nullptr, nullptr);
        {
            std::lock_guard lock(clearMutex);
            clearHr        = localHr;
            clearCompleted = true;
        }
        clearCv.notify_all();
    });
    auto joinClearThread = wil::scope_exit([&]() noexcept
    {
        if (clearThread.joinable())
        {
            clearThread.join();
        }
    });

    const bool clearReturnedWhileHostBlocked = [&]() noexcept
    {
        std::unique_lock lock(clearMutex);
        return clearCv.wait_for(lock, std::chrono::seconds(2), [&]() noexcept { return clearCompleted; });
    }();

    state.Require(clearReturnedWhileHostBlocked,
                  L"Google Drive navigation callback drain: SetCallback(nullptr, nullptr) should return without waiting for blocked host UI.");

    host->ReleaseShowConnectionManager();
    clearThread.join();
    staleCommandThread.join();

    state.Require(SUCCEEDED(clearHr),
                  std::format(L"Google Drive navigation callback drain: clearing callback failed. hr=0x{:08X}", static_cast<unsigned long>(clearHr)));
    state.Require(staleCommandHr == HRESULT_FROM_WIN32(ERROR_CANCELLED),
                  std::format(L"Google Drive navigation callback drain: stale ExecuteMenuCommand should self-drop after clear. hr=0x{:08X}",
                              static_cast<unsigned long>(staleCommandHr)));
    state.Require(callbackProbe.GetCallCount() == 0u, L"Google Drive navigation callback drain: callback fired after SetCallback(nullptr, nullptr) returned.");
    state.Require(callbackProbe.GetLastPath().empty(), L"Google Drive navigation callback drain: stale callback path should stay empty after clear.");

    return state.failure.empty();
});

SelfTest::RunCase(options,
                  suite,
                  L"google_drive_cleared_client_id_requires_configuration",
                  [&](SelfTest::CaseState& state) noexcept
{
    CreatedFileSystemInstance created{};
    const HRESULT createHr = TryCreateFileSystemInstance(kBuiltinGoogleDriveFileSystemId, {}, created);
    state.Require(SUCCEEDED(createHr) && created.fileSystem,
                  std::format(L"Google Drive clientId gate: failed to create filesystem instance. hr=0x{:08X}", static_cast<unsigned long>(createHr)));
    if (FAILED(createHr) || ! created.fileSystem)
    {
        return false;
    }

    wil::com_ptr<IInformations> info;
    state.Require(CreateInformations(created.fileSystem, info), L"Google Drive clientId gate: missing IInformations.");
    if (! info)
    {
        return false;
    }

    HRESULT hr = info->SetConfiguration("{\"defaultClientId\":\"\"}");
    state.Require(SUCCEEDED(hr), std::format(L"Google Drive clientId gate: SetConfiguration failed. hr=0x{:08X}", static_cast<unsigned long>(hr)));
    if (FAILED(hr))
    {
        return false;
    }

    const std::wstring id = MakeGuidText();
    state.Require(! id.empty(), L"Google Drive clientId gate: failed to generate a connection profile id.");
    if (id.empty())
    {
        return false;
    }

    Common::Settings::ConnectionProfile profile;
    profile.id                  = id;
    profile.name                = std::format(L"GDriveMissingClientId{}", id);
    profile.pluginId            = std::wstring(kBuiltinGoogleDriveFileSystemId);
    profile.host                = L"";
    profile.initialPath         = L"/";
    profile.userName            = L"user@example.invalid";
    profile.authMode            = Common::Settings::ConnectionAuthMode::OAuth2Pkce;
    profile.savePassword        = false;
    profile.requireWindowsHello = false;

    const bool hadConnectionsSettings = g_settings.connections.has_value();
    if (! hadConnectionsSettings)
    {
        g_settings.connections.emplace();
    }

    g_settings.connections->items.push_back(profile);
    const auto restoreSettings = wil::scope_exit([&] noexcept
    {
        if (! g_settings.connections)
        {
            return;
        }

        auto& items = g_settings.connections->items;
        items.erase(std::remove_if(items.begin(), items.end(), [&](const Common::Settings::ConnectionProfile& item) noexcept { return item.id == id; }),
                    items.end());

        if (! hadConnectionsSettings)
        {
            g_settings.connections.reset();
        }
    });

    wil::com_ptr<IHostAlerts> hostAlerts;
    static_cast<void>(GetHostServices()->QueryInterface(IID_PPV_ARGS(hostAlerts.put())));
    const auto clearAlert = wil::scope_exit([&] noexcept
    {
        if (hostAlerts)
        {
            static_cast<void>(hostAlerts->ClearAlert(HOST_ALERT_SCOPE_APPLICATION, nullptr));
        }
    });

    const std::wstring connectionRoot = std::format(L"/@conn:{}/", profile.name);
    wil::com_ptr<IFilesInformation> filesInformation;
    hr = created.fileSystem->ReadDirectoryInfo(connectionRoot.c_str(), filesInformation.put());
    if (SkipIfHostConnectionUiUnavailable(state, hr, L"Google Drive clientId gate"))
    {
        return true;
    }
    state.Require(hr == HRESULT_FROM_WIN32(ERROR_BAD_CONFIGURATION),
                  std::format(L"Google Drive clientId gate: expected ERROR_BAD_CONFIGURATION. hr=0x{:08X}", static_cast<unsigned long>(hr)));
    state.Require(! filesInformation, L"Google Drive clientId gate: files information should not be produced on configuration failure.");

    return state.failure.empty();
});

SelfTest::RunCase(options,
                  suite,
                  L"google_drive_connection_requires_refresh_token",
                  [&](SelfTest::CaseState& state) noexcept
{
    CreatedFileSystemInstance created{};
    const HRESULT createHr = TryCreateFileSystemInstance(kBuiltinGoogleDriveFileSystemId, {}, created);
    state.Require(SUCCEEDED(createHr) && created.fileSystem,
                  std::format(L"Google Drive refresh-token gate: failed to create filesystem instance. hr=0x{:08X}", static_cast<unsigned long>(createHr)));
    if (FAILED(createHr) || ! created.fileSystem)
    {
        return false;
    }

    wil::com_ptr<IInformations> info;
    state.Require(CreateInformations(created.fileSystem, info), L"Google Drive refresh-token gate: missing IInformations.");
    if (! info)
    {
        return false;
    }

    HRESULT hr = info->SetConfiguration("{\"defaultClientId\":\"selftest-client-id.apps.googleusercontent.com\"}");
    state.Require(SUCCEEDED(hr), std::format(L"Google Drive refresh-token gate: SetConfiguration failed. hr=0x{:08X}", static_cast<unsigned long>(hr)));
    if (FAILED(hr))
    {
        return false;
    }

    const std::wstring id = MakeGuidText();
    state.Require(! id.empty(), L"Google Drive refresh-token gate: failed to generate a connection profile id.");
    if (id.empty())
    {
        return false;
    }

    Common::Settings::ConnectionProfile profile;
    profile.id                  = id;
    profile.name                = std::format(L"GDriveSelfTest{}", id);
    profile.pluginId            = std::wstring(kBuiltinGoogleDriveFileSystemId);
    profile.host                = L"";
    profile.initialPath         = L"/";
    profile.userName            = L"user@example.invalid";
    profile.authMode            = Common::Settings::ConnectionAuthMode::OAuth2Pkce;
    profile.savePassword        = false;
    profile.requireWindowsHello = false;

    const bool hadConnectionsSettings = g_settings.connections.has_value();
    if (! hadConnectionsSettings)
    {
        g_settings.connections.emplace();
    }

    g_settings.connections->items.push_back(profile);
    const auto restoreSettings = wil::scope_exit([&] noexcept
    {
        if (! g_settings.connections)
        {
            return;
        }

        auto& items = g_settings.connections->items;
        items.erase(std::remove_if(items.begin(), items.end(), [&](const Common::Settings::ConnectionProfile& item) noexcept { return item.id == id; }),
                    items.end());

        if (! hadConnectionsSettings)
        {
            g_settings.connections.reset();
        }
    });

    const std::wstring connectionRoot = std::format(L"/@conn:{}/", profile.name);
    wil::com_ptr<IFilesInformation> filesInformation;
    hr = created.fileSystem->ReadDirectoryInfo(connectionRoot.c_str(), filesInformation.put());
    if (SkipIfHostConnectionUiUnavailable(state, hr, L"Google Drive refresh-token gate"))
    {
        return true;
    }
    state.Require(
        hr == HRESULT_FROM_WIN32(ERROR_NOT_FOUND),
        std::format(L"Google Drive refresh-token gate: expected ERROR_NOT_FOUND when no refresh token is saved. hr=0x{:08X}", static_cast<unsigned long>(hr)));

    return state.failure.empty();
});

SelfTest::RunCase(options,
                  suite,
                  L"onedrive_personal_cleared_client_id_requires_configuration",
                  [&](SelfTest::CaseState& state) noexcept
{
    CreatedFileSystemInstance created{};
    const HRESULT createHr = TryCreateFileSystemInstance(kBuiltinOneDrivePersonalFileSystemId, {}, created);
    state.Require(SUCCEEDED(createHr) && created.fileSystem,
                  std::format(L"OneDrive Personal plugin: failed to create filesystem instance. hr=0x{:08X}", static_cast<unsigned long>(createHr)));
    if (FAILED(createHr) || ! created.fileSystem)
    {
        return false;
    }

    wil::com_ptr<IInformations> info;
    state.Require(CreateInformations(created.fileSystem, info), L"OneDrive Personal plugin: missing IInformations.");
    if (! info)
    {
        return false;
    }

    const char* schema = nullptr;
    HRESULT hr         = info->GetConfigurationSchema(&schema);
    state.Require(SUCCEEDED(hr) && schema,
                  std::format(L"OneDrive Personal plugin: GetConfigurationSchema failed. hr=0x{:08X}", static_cast<unsigned long>(hr)));
    if (FAILED(hr) || ! schema)
    {
        return false;
    }

    const std::string_view schemaView(schema);
    state.Require(schemaView.find("\"clientId\"") != std::string_view::npos, L"OneDrive Personal plugin: schema missing clientId.");
    state.Require(schemaView.find("90cdea53-7c21-48b0-959e-b4024209027b") != std::string_view::npos,
                  L"OneDrive Personal plugin: schema missing built-in default clientId.");
    state.Require(schemaView.find("\"x-ui-hidden\"") != std::string_view::npos, L"OneDrive Personal plugin: schema missing x-ui-hidden metadata for clientId.");

    hr = info->SetConfiguration("{\"clientId\":\"\"}");
    state.Require(SUCCEEDED(hr),
                  std::format(L"OneDrive Personal plugin: SetConfiguration with explicit empty clientId failed. hr=0x{:08X}", static_cast<unsigned long>(hr)));
    if (FAILED(hr))
    {
        return false;
    }

    const std::wstring id = MakeGuidText();
    state.Require(! id.empty(), L"OneDrive Personal clientId gate: failed to generate a connection profile id.");
    if (id.empty())
    {
        return false;
    }

    Common::Settings::ConnectionProfile profile;
    profile.id                  = id;
    profile.name                = std::format(L"OneDriveMissingClientId{}", id);
    profile.pluginId            = std::wstring(kBuiltinOneDrivePersonalFileSystemId);
    profile.host                = L"";
    profile.initialPath         = L"/";
    profile.userName            = L"user@example.invalid";
    profile.authMode            = Common::Settings::ConnectionAuthMode::OAuth2Pkce;
    profile.savePassword        = false;
    profile.requireWindowsHello = false;

    const bool hadConnectionsSettings = g_settings.connections.has_value();
    if (! hadConnectionsSettings)
    {
        g_settings.connections.emplace();
    }

    g_settings.connections->items.push_back(profile);
    const auto restoreSettings = wil::scope_exit([&] noexcept
    {
        if (! g_settings.connections)
        {
            return;
        }

        auto& items = g_settings.connections->items;
        items.erase(std::remove_if(items.begin(), items.end(), [&](const Common::Settings::ConnectionProfile& item) noexcept { return item.id == id; }),
                    items.end());

        if (! hadConnectionsSettings)
        {
            g_settings.connections.reset();
        }
    });

    wil::com_ptr<IHostAlerts> hostAlerts;
    static_cast<void>(GetHostServices()->QueryInterface(IID_PPV_ARGS(hostAlerts.put())));
    const auto clearAlert = wil::scope_exit([&] noexcept
    {
        if (hostAlerts)
        {
            static_cast<void>(hostAlerts->ClearAlert(HOST_ALERT_SCOPE_APPLICATION, nullptr));
        }
    });

    const std::wstring connectionRoot = std::format(L"/@conn:{}/", profile.name);
    wil::com_ptr<IFilesInformation> filesInformation;
    hr = created.fileSystem->ReadDirectoryInfo(connectionRoot.c_str(), filesInformation.put());
    if (SkipIfHostConnectionUiUnavailable(state, hr, L"OneDrive Personal clientId gate"))
    {
        return true;
    }
    state.Require(hr == HRESULT_FROM_WIN32(ERROR_BAD_CONFIGURATION),
                  std::format(L"OneDrive Personal clientId gate: expected ERROR_BAD_CONFIGURATION. hr=0x{:08X}", static_cast<unsigned long>(hr)));
    state.Require(! filesInformation, L"OneDrive Personal clientId gate: files information should not be produced on configuration failure.");

    return state.failure.empty();
});

SelfTest::RunCase(options,
                  suite,
                  L"unique",
                  [&](SelfTest::CaseState& state) noexcept
{
    // Case: Unique files/dirs selected; identical excluded by default.
    if (const auto foldersOpt = CreateCaseFolders(root, L"unique"))
    {
        const auto& folders = foldersOpt.value();
        state.Require(SelfTest::WriteTextFile(folders.left / L"only_left.txt", "L"), L"Failed to create only_left.txt (left).");
        state.Require(SelfTest::WriteTextFile(folders.right / L"only_right.txt", "R"), L"Failed to create only_right.txt (right).");
        state.Require(SelfTest::EnsureDirectory(folders.left / L"only_left_dir"), L"Failed to create only_left_dir (left).");
        state.Require(SelfTest::WriteTextFile(folders.left / L"same.txt", "S"), L"Failed to create same.txt (left).");
        state.Require(SelfTest::WriteTextFile(folders.right / L"same.txt", "S"), L"Failed to create same.txt (right).");

        AppendCompareSelfTestTraceLine(L"Case: unique: computing decision");
        auto decision = ComputeRootDecision(baseFs, folders, Common::Settings::CompareDirectoriesSettings{}, state);
        AppendCompareSelfTestTraceLine(L"Case: unique: decision returned");
        if (decision)
        {
            {
                const auto* item = FindItem(*decision, L"only_left.txt");
                state.Require(item != nullptr, L"only_left.txt missing from decision.");
                if (item)
                {
                    state.Require(item->isDifferent, L"only_left.txt expected isDifferent.");
                    state.Require(item->selectLeft && ! item->selectRight, L"only_left.txt expected selectLeft only.");
                    state.Require(HasFlag(item->differenceMask, CompareDirectoriesDiffBit::OnlyInLeft), L"only_left.txt expected differenceMask=OnlyInLeft.");
                }
            }
            {
                const auto* item = FindItem(*decision, L"only_right.txt");
                state.Require(item != nullptr, L"only_right.txt missing from decision.");
                if (item)
                {
                    state.Require(item->isDifferent, L"only_right.txt expected isDifferent.");
                    state.Require(! item->selectLeft && item->selectRight, L"only_right.txt expected selectRight only.");
                    state.Require(HasFlag(item->differenceMask, CompareDirectoriesDiffBit::OnlyInRight),
                                  L"only_right.txt expected differenceMask=OnlyInRight.");
                }
            }
            {
                const auto* item = FindItem(*decision, L"only_left_dir");
                state.Require(item != nullptr, L"only_left_dir missing from decision.");
                if (item)
                {
                    state.Require(item->isDirectory, L"only_left_dir expected isDirectory.");
                    state.Require(item->isDifferent, L"only_left_dir expected isDifferent.");
                    state.Require(item->selectLeft && ! item->selectRight, L"only_left_dir expected selectLeft only.");
                    state.Require(HasFlag(item->differenceMask, CompareDirectoriesDiffBit::OnlyInLeft), L"only_left_dir expected differenceMask=OnlyInLeft.");
                }
            }
            {
                const auto* item = FindItem(*decision, L"same.txt");
                state.Require(item == nullptr, L"same.txt expected elided from decision in differences-only mode.");
            }

            auto session =
                std::make_shared<CompareDirectoriesSession>(baseFs, baseFs, folders.left, folders.right, Common::Settings::CompareDirectoriesSettings{});
            const auto fsLeft  = CreateCompareDirectoriesFileSystem(ComparePane::Left, session);
            const auto fsRight = CreateCompareDirectoriesFileSystem(ComparePane::Right, session);

            const auto scanTimeout = std::chrono::milliseconds(SelfTest::ScaleTimeout(5'000));
            state.Require(StartScanAndWaitForIdle(session, scanTimeout), L"unique: scan did not go idle.");
            state.Require(DrainPendingSubdirUpdates(session, 256), L"unique: subdir updates did not drain.");

            const auto leftNames  = EnumerateDirectoryNames(fsLeft, folders.left, state);
            const auto rightNames = EnumerateDirectoryNames(fsRight, folders.right, state);
            AppendCompareSelfTestTraceLine(L"Case: unique: enumeration done");

            state.Require(ContainsName(leftNames, L"only_left.txt"), L"only_left.txt expected in left enumeration.");
            state.Require(! ContainsName(leftNames, L"only_right.txt"), L"only_right.txt unexpected in left enumeration.");
            state.Require(! ContainsName(leftNames, L"same.txt"), L"same.txt expected excluded in left enumeration.");

            state.Require(ContainsName(rightNames, L"only_right.txt"), L"only_right.txt expected in right enumeration.");
            state.Require(! ContainsName(rightNames, L"only_left.txt"), L"only_left.txt unexpected in right enumeration.");
            state.Require(! ContainsName(rightNames, L"same.txt"), L"same.txt expected excluded in right enumeration.");

            AppendCompareSelfTestTraceLine(L"Case: unique: done");
        }
    }
    else
    {
        state.Require(false, L"Failed to create case folders: unique.");
    }

    return state.failure.empty();
});

SelfTest::RunCase(options,
                  suite,
                  L"typemismatch",
                  [&](SelfTest::CaseState& state) noexcept
{
    // Case: File vs directory mismatch selects both sides.
    if (const auto foldersOpt = CreateCaseFolders(root, L"typemismatch"))
    {
        const auto& folders = foldersOpt.value();
        state.Require(SelfTest::WriteTextFile(folders.left / L"mix", "F"), L"Failed to create mix file (left).");
        state.Require(SelfTest::EnsureDirectory(folders.right / L"mix"), L"Failed to create mix directory (right).");

        auto decision = ComputeRootDecision(baseFs, folders, Common::Settings::CompareDirectoriesSettings{}, state);
        if (decision)
        {
            const auto* item = FindItem(*decision, L"mix");
            state.Require(item != nullptr, L"mix missing from decision.");
            if (item)
            {
                state.Require(item->isDifferent, L"mix expected isDifferent on type mismatch.");
                state.Require(item->selectLeft && item->selectRight, L"mix expected select both on type mismatch.");
                state.Require(HasFlag(item->differenceMask, CompareDirectoriesDiffBit::TypeMismatch), L"mix expected differenceMask=TypeMismatch.");
            }
        }
    }
    else
    {
        state.Require(false, L"Failed to create case folders: typemismatch.");
    }

    return state.failure.empty();
});

SelfTest::RunCase(options,
                  suite,
                  L"size",
                  [&](SelfTest::CaseState& state) noexcept
{
    // Case: Size compare selects bigger file.
    if (const auto foldersOpt = CreateCaseFolders(root, L"size"))
    {
        const auto& folders = foldersOpt.value();
        state.Require(WriteFileFill(folders.left / L"a.bin", 'A', 200), L"Failed to create a.bin (left).");
        state.Require(WriteFileFill(folders.right / L"a.bin", 'B', 100), L"Failed to create a.bin (right).");

        Common::Settings::CompareDirectoriesSettings settings{};
        settings.compareSize = true;

        auto decision = ComputeRootDecision(baseFs, folders, settings, state);
        if (decision)
        {
            const auto* item = FindItem(*decision, L"a.bin");
            state.Require(item != nullptr, L"a.bin missing from decision.");
            if (item)
            {
                state.Require(item->isDifferent, L"a.bin expected isDifferent with compareSize.");
                state.Require(item->selectLeft && ! item->selectRight, L"a.bin expected selectLeft only when left is bigger.");
                state.Require(HasFlag(item->differenceMask, CompareDirectoriesDiffBit::Size), L"a.bin expected differenceMask=Size.");
            }
        }
    }
    else
    {
        state.Require(false, L"Failed to create case folders: size.");
    }

    return state.failure.empty();
});

SelfTest::RunCase(options,
                  suite,
                  L"time",
                  [&](SelfTest::CaseState& state) noexcept
{
    // Case: Date/time compare selects newer file.
    if (const auto foldersOpt = CreateCaseFolders(root, L"time"))
    {
        const auto& folders = foldersOpt.value();
        state.Require(SelfTest::WriteTextFile(folders.left / L"a.txt", "T"), L"Failed to create a.txt (left).");
        state.Require(SelfTest::WriteTextFile(folders.right / L"a.txt", "T"), L"Failed to create a.txt (right).");

        FILETIME now{};
        ::GetSystemTimeAsFileTime(&now);
        ULARGE_INTEGER newer{};
        newer.LowPart  = now.dwLowDateTime;
        newer.HighPart = now.dwHighDateTime;
        newer.QuadPart += 60ull * 10'000'000ull;

        FILETIME leftFt{};
        leftFt.dwLowDateTime  = newer.LowPart;
        leftFt.dwHighDateTime = newer.HighPart;

        state.Require(SetFileLastWriteTime(folders.left / L"a.txt", leftFt), L"Failed to set a.txt last write time (left).");
        state.Require(SetFileLastWriteTime(folders.right / L"a.txt", now), L"Failed to set a.txt last write time (right).");

        Common::Settings::CompareDirectoriesSettings settings{};
        settings.compareDateTime = true;

        auto decision = ComputeRootDecision(baseFs, folders, settings, state);
        if (decision)
        {
            const auto* item = FindItem(*decision, L"a.txt");
            state.Require(item != nullptr, L"a.txt missing from decision.");
            if (item)
            {
                state.Require(item->isDifferent, L"a.txt expected isDifferent with compareDateTime.");
                state.Require(item->selectLeft && ! item->selectRight, L"a.txt expected selectLeft only when left is newer.");
                state.Require(HasFlag(item->differenceMask, CompareDirectoriesDiffBit::DateTime), L"a.txt expected differenceMask=DateTime.");
            }
        }
    }
    else
    {
        state.Require(false, L"Failed to create case folders: time.");
    }

    return state.failure.empty();
});

SelfTest::RunCase(options,
                  suite,
                  L"attributes",
                  [&](SelfTest::CaseState& state) noexcept
{
    // Case: Attribute compare selects both sides.
    if (const auto foldersOpt = CreateCaseFolders(root, L"attributes"))
    {
        const auto& folders = foldersOpt.value();
        state.Require(SelfTest::WriteTextFile(folders.left / L"a.txt", "A"), L"Failed to create a.txt (left).");
        state.Require(SelfTest::WriteTextFile(folders.right / L"a.txt", "A"), L"Failed to create a.txt (right).");

        const std::filesystem::path leftPath = folders.left / L"a.txt";
        const DWORD leftAttrs                = ::GetFileAttributesW(leftPath.c_str());
        state.Require(leftAttrs != INVALID_FILE_ATTRIBUTES, L"GetFileAttributesW failed for a.txt (left).");
        if (leftAttrs != INVALID_FILE_ATTRIBUTES)
        {
            state.Require(::SetFileAttributesW(leftPath.c_str(), leftAttrs | FILE_ATTRIBUTE_HIDDEN) != 0, L"SetFileAttributesW failed for a.txt (left).");
        }

        Common::Settings::CompareDirectoriesSettings settings{};
        settings.compareAttributes = true;

        auto decision = ComputeRootDecision(baseFs, folders, settings, state);
        if (decision)
        {
            const auto* item = FindItem(*decision, L"a.txt");
            state.Require(item != nullptr, L"a.txt missing from decision.");
            if (item)
            {
                state.Require(item->isDifferent, L"a.txt expected isDifferent with compareAttributes.");
                state.Require(item->selectLeft && item->selectRight, L"a.txt expected select both when attributes differ.");
                state.Require(HasFlag(item->differenceMask, CompareDirectoriesDiffBit::Attributes), L"a.txt expected differenceMask=Attributes.");
            }
        }
    }
    else
    {
        state.Require(false, L"Failed to create case folders: attributes.");
    }

    return state.failure.empty();
});

SelfTest::RunCase(options,
                  suite,
                  L"content",
                  [&](SelfTest::CaseState& state) noexcept
{
    // Case: Content compare selects both sides.
    if (const auto foldersOpt = CreateCaseFolders(root, L"content"))
    {
        const auto& folders = foldersOpt.value();
        state.Require(WriteFileFill(folders.left / L"a.bin", 'X', 64), L"Failed to create a.bin (left).");
        state.Require(WriteFileFill(folders.right / L"a.bin", 'Y', 64), L"Failed to create a.bin (right).");

        Common::Settings::CompareDirectoriesSettings settings{};
        settings.compareContent = true;

        auto session  = std::make_shared<CompareDirectoriesSession>(baseFs, baseFs, folders.left, folders.right, settings);
        auto decision = WaitForContentCompare(session, std::filesystem::path{}, L"a.bin", state);
        if (decision)
        {
            const auto* item = FindItem(*decision, L"a.bin");
            state.Require(item != nullptr, L"a.bin missing from decision.");
            if (item)
            {
                state.Require(item->isDifferent, L"a.bin expected isDifferent with compareContent.");
                state.Require(item->selectLeft && item->selectRight, L"a.bin expected select both when content differs.");
                state.Require(HasFlag(item->differenceMask, CompareDirectoriesDiffBit::Content), L"a.bin expected differenceMask=Content.");
                state.Require(! HasFlag(item->differenceMask, CompareDirectoriesDiffBit::ContentPending),
                              L"a.bin expected ContentPending cleared after compare completes.");
            }
        }
    }
    else
    {
        state.Require(false, L"Failed to create case folders: content.");
    }

    return state.failure.empty();
});

SelfTest::RunCase(options,
                  suite,
                  L"content_dual_io",
                  [&](SelfTest::CaseState& state) noexcept
{
    // Case: Content compare uses the correct per-pane IFileSystemIO (dual-IO regression guard).
    if (const auto foldersOpt = CreateCaseFolders(root, L"content_dual_io"))
    {
        const auto& folders = foldersOpt.value();
        state.Require(SelfTest::WriteTextFile(folders.left / L"a.txt", "AAAA"), L"Failed to create a.txt (left).");
        state.Require(SelfTest::WriteTextFile(folders.right / L"a.txt", "BBBB"), L"Failed to create a.txt (right).");

        const wil::com_ptr<IFileSystem> leftFs  = CreatePluginPathMappedRootFileSystem(baseFs, folders.left);
        const wil::com_ptr<IFileSystem> rightFs = CreatePluginPathMappedRootFileSystem(baseFs, folders.right);
        state.Require(static_cast<bool>(leftFs), L"Failed to create left mapped filesystem.");
        state.Require(static_cast<bool>(rightFs), L"Failed to create right mapped filesystem.");

        Common::Settings::CompareDirectoriesSettings settings{};
        settings.compareContent = true;

        const std::filesystem::path pluginRoot(L"/");
        auto session = std::make_shared<CompareDirectoriesSession>(leftFs, rightFs, pluginRoot, pluginRoot, settings);
        state.Require(session->IsContentCompareSupported(), L"content_dual_io: content compare should be supported.");

        auto decision = WaitForContentCompare(session, std::filesystem::path{}, L"a.txt", state);
        if (decision)
        {
            const auto* item = FindItem(*decision, L"a.txt");
            state.Require(item != nullptr, L"a.txt missing from decision.");
            if (item)
            {
                state.Require(item->isDifferent, L"a.txt expected isDifferent with compareContent.");
                state.Require(item->selectLeft && item->selectRight, L"a.txt expected select both when content differs.");
                state.Require(HasFlag(item->differenceMask, CompareDirectoriesDiffBit::Content), L"a.txt expected differenceMask=Content.");
                state.Require(! HasFlag(item->differenceMask, CompareDirectoriesDiffBit::ContentPending),
                              L"a.txt expected ContentPending cleared after compare completes.");
            }
        }
    }
    else
    {
        state.Require(false, L"Failed to create case folders: content_dual_io.");
    }

    return state.failure.empty();
});

SelfTest::RunCase(options,
                  suite,
                  L"content_equal_size_equal_mtime_differs",
                  [&](SelfTest::CaseState& state) noexcept
{
    // Case: equal size and equal mtime still require byte comparison when compareContent=true.
    if (const auto foldersOpt = CreateCaseFolders(root, L"content_equal_size_equal_mtime_differs"))
    {
        const auto& folders = foldersOpt.value();
        state.Require(SelfTest::WriteTextFile(folders.left / L"diff.bin", "ABCD"), L"content equal metadata: failed to write left diff.bin.");
        state.Require(SelfTest::WriteTextFile(folders.right / L"diff.bin", "WXYZ"), L"content equal metadata: failed to write right diff.bin.");
        state.Require(SelfTest::WriteTextFile(folders.left / L"same.bin", "SAME"), L"content equal metadata: failed to write left same.bin.");
        state.Require(SelfTest::WriteTextFile(folders.right / L"same.bin", "SAME"), L"content equal metadata: failed to write right same.bin.");

        FILETIME fixedTime{};
        ::GetSystemTimeAsFileTime(&fixedTime);
        state.Require(SetFileLastWriteTime(folders.left / L"diff.bin", fixedTime), L"content equal metadata: failed to set left diff.bin mtime.");
        state.Require(SetFileLastWriteTime(folders.right / L"diff.bin", fixedTime), L"content equal metadata: failed to set right diff.bin mtime.");
        state.Require(SetFileLastWriteTime(folders.left / L"same.bin", fixedTime), L"content equal metadata: failed to set left same.bin mtime.");
        state.Require(SetFileLastWriteTime(folders.right / L"same.bin", fixedTime), L"content equal metadata: failed to set right same.bin mtime.");

        Common::Settings::CompareDirectoriesSettings settings{};
        settings.compareContent     = true;
        settings.keepIdenticalItems = true;

        auto session      = std::make_shared<CompareDirectoriesSession>(baseFs, baseFs, folders.left, folders.right, settings);
        auto diffDecision = WaitForContentCompare(session, std::filesystem::path{}, L"diff.bin", state);
        if (diffDecision)
        {
            const auto* item = FindItem(*diffDecision, L"diff.bin");
            state.Require(item != nullptr, L"content equal metadata: diff.bin missing.");
            if (item)
            {
                state.Require(item->isDifferent, L"content equal metadata: diff.bin must be different despite equal size/mtime.");
                state.Require(item->selectLeft && item->selectRight, L"content equal metadata: diff.bin must select both sides.");
                state.Require(HasFlag(item->differenceMask, CompareDirectoriesDiffBit::Content), L"content equal metadata: diff.bin missing Content bit.");
            }
        }

        auto sameDecision = WaitForContentCompare(session, std::filesystem::path{}, L"same.bin", state);
        if (sameDecision)
        {
            const auto* item = FindItem(*sameDecision, L"same.bin");
            state.Require(item != nullptr, L"content equal metadata: same.bin missing with keepIdenticalItems=true.");
            if (item)
            {
                state.Require(! item->isDifferent, L"content equal metadata: same.bin must remain identical after byte compare.");
                state.Require(item->differenceMask == 0u, L"content equal metadata: same.bin expected no difference bits.");
                state.Require(! item->selectLeft && ! item->selectRight, L"content equal metadata: same.bin expected no selection.");
            }
        }
    }
    else
    {
        state.Require(false, L"Failed to create case folders: content_equal_size_equal_mtime_differs.");
    }

    return state.failure.empty();
});

SelfTest::RunCase(options,
                  suite,
                  L"content_unknown_size_streaming_compare",
                  [&](SelfTest::CaseState& state) noexcept
{
    // Case: unknown-size readers use streaming EOF comparison for equal, prefix, and mismatch files.
    if (const auto foldersOpt = CreateCaseFolders(root, L"content_unknown_size_streaming_compare"))
    {
        const auto& folders = foldersOpt.value();
        state.Require(WriteFileFill(folders.left / L"empty.bin", 'E', 0), L"unknown size: failed to write left empty.bin.");
        state.Require(WriteFileFill(folders.right / L"empty.bin", 'E', 0), L"unknown size: failed to write right empty.bin.");
        state.Require(SelfTest::WriteTextFile(folders.left / L"equal.bin", "abcdef"), L"unknown size: failed to write left equal.bin.");
        state.Require(SelfTest::WriteTextFile(folders.right / L"equal.bin", "abcdef"), L"unknown size: failed to write right equal.bin.");
        state.Require(SelfTest::WriteTextFile(folders.left / L"left_prefix.bin", "abc"), L"unknown size: failed to write left_prefix left.");
        state.Require(SelfTest::WriteTextFile(folders.right / L"left_prefix.bin", "abcdef"), L"unknown size: failed to write left_prefix right.");
        state.Require(SelfTest::WriteTextFile(folders.left / L"right_prefix.bin", "abcdef"), L"unknown size: failed to write right_prefix left.");
        state.Require(SelfTest::WriteTextFile(folders.right / L"right_prefix.bin", "abc"), L"unknown size: failed to write right_prefix right.");
        state.Require(SelfTest::WriteTextFile(folders.left / L"mid_mismatch.bin", "abcXef"), L"unknown size: failed to write mid_mismatch left.");
        state.Require(SelfTest::WriteTextFile(folders.right / L"mid_mismatch.bin", "abcYef"), L"unknown size: failed to write mid_mismatch right.");

        Common::Settings::CompareDirectoriesSettings settings{};
        settings.compareContent     = true;
        settings.keepIdenticalItems = true;

        wil::com_ptr<IFileSystem> wrapped = CreateShortReadFileSystem(baseFs, folders.left.parent_path(), 2u, 0u, true);
        state.Require(static_cast<bool>(wrapped), L"unknown size: failed to create unknown-size wrapper.");

        auto session =
            std::make_shared<CompareDirectoriesSession>(wrapped ? wrapped : baseFs, wrapped ? wrapped : baseFs, folders.left, folders.right, settings);

        const auto assertItem = [&](std::wstring_view name, bool expectedDifferent) noexcept
        {
            auto decision = WaitForContentCompare(session, std::filesystem::path{}, name, state);
            if (! decision)
            {
                return;
            }
            const auto* item = FindItem(*decision, name);
            state.Require(item != nullptr, std::format(L"unknown size: {} missing.", name));
            if (! item)
            {
                return;
            }
            state.Require(item->isDifferent == expectedDifferent, std::format(L"unknown size: {} different state mismatch.", name));
            state.Require(HasFlag(item->differenceMask, CompareDirectoriesDiffBit::Content) == expectedDifferent,
                          std::format(L"unknown size: {} Content bit mismatch.", name));
            state.Require(! HasFlag(item->differenceMask, CompareDirectoriesDiffBit::ContentPending),
                          std::format(L"unknown size: {} ContentPending must be cleared.", name));
            state.Require((item->selectLeft && item->selectRight) == expectedDifferent, std::format(L"unknown size: {} selection mismatch.", name));
        };

        assertItem(L"empty.bin", false);
        assertItem(L"equal.bin", false);
        assertItem(L"left_prefix.bin", true);
        assertItem(L"right_prefix.bin", true);
        assertItem(L"mid_mismatch.bin", true);
    }
    else
    {
        state.Require(false, L"Failed to create case folders: content_unknown_size_streaming_compare.");
    }

    return state.failure.empty();
});

SelfTest::RunCase(options,
                  suite,
                  L"content_no_io_disables_compareContent",
                  [&](SelfTest::CaseState& state) noexcept
{
    // Case: compareContent is treated as disabled when either side lacks IFileSystemIO.
    if (const auto foldersOpt = CreateCaseFolders(root, L"content_no_io_disables_compareContent"))
    {
        const auto& folders = foldersOpt.value();
        state.Require(SelfTest::WriteTextFile(folders.left / L"a.txt", "AAAA"), L"Failed to create a.txt (left).");
        state.Require(SelfTest::WriteTextFile(folders.right / L"a.txt", "BBBB"), L"Failed to create a.txt (right).");

        const wil::com_ptr<IFileSystem> leftFs  = CreatePluginPathMappedRootFileSystem(baseFs, folders.left);
        const wil::com_ptr<IFileSystem> rightFs = CreatePluginPathMappedRootFileSystemNoIO(baseFs, folders.right);
        state.Require(static_cast<bool>(leftFs), L"Failed to create left mapped filesystem.");
        state.Require(static_cast<bool>(rightFs), L"Failed to create right mapped filesystem (no IO).");

        Common::Settings::CompareDirectoriesSettings settings{};
        settings.compareContent     = true;
        settings.keepIdenticalItems = true;

        const std::filesystem::path pluginRoot(L"/");
        const auto session = std::make_shared<CompareDirectoriesSession>(leftFs, rightFs, pluginRoot, pluginRoot, settings);
        state.Require(! session->IsContentCompareSupported(), L"content_no_io: content compare should be unsupported.");

        const auto decision = session->GetOrComputeDecision(std::filesystem::path{});
        state.Require(static_cast<bool>(decision), L"content_no_io: decision is null.");
        if (decision)
        {
            state.Require(SUCCEEDED(decision->hr), L"content_no_io: decision hr is failure.");

            const auto* item = FindItem(*decision, L"a.txt");
            state.Require(item != nullptr, L"content_no_io: a.txt missing from decision.");
            if (item)
            {
                state.Require(! item->isDifferent, L"content_no_io: a.txt expected identical (content criterion disabled).");
                state.Require(item->differenceMask == 0u, L"content_no_io: a.txt expected differenceMask=0.");
                state.Require(! HasFlag(item->differenceMask, CompareDirectoriesDiffBit::ContentPending),
                              L"content_no_io: a.txt expected no ContentPending when unsupported.");
            }
        }
    }
    else
    {
        state.Require(false, L"Failed to create case folders: content_no_io_disables_compareContent.");
    }

    return state.failure.empty();
});

SelfTest::RunCase(options,
                  suite,
                  L"content_size_mismatch_no_pending",
                  [&](SelfTest::CaseState& state) noexcept
{
    // Case: Content compare with different sizes does not mark ContentPending.
    if (const auto foldersOpt = CreateCaseFolders(root, L"content_size_mismatch_no_pending"))
    {
        const auto& folders = foldersOpt.value();
        state.Require(WriteFileFill(folders.left / L"a.bin", 'X', 64), L"Failed to create a.bin (left).");
        state.Require(WriteFileFill(folders.right / L"a.bin", 'X', 32), L"Failed to create a.bin (right).");

        Common::Settings::CompareDirectoriesSettings settings{};
        settings.compareContent = true;

        auto decision = ComputeRootDecision(baseFs, folders, settings, state);
        if (decision)
        {
            const auto* item = FindItem(*decision, L"a.bin");
            state.Require(item != nullptr, L"a.bin missing from decision.");
            if (item)
            {
                state.Require(item->isDifferent, L"a.bin expected isDifferent with compareContent and size mismatch.");
                state.Require(item->selectLeft && item->selectRight, L"a.bin expected select both when compareContent and sizes differ.");
                state.Require(HasFlag(item->differenceMask, CompareDirectoriesDiffBit::Content), L"a.bin expected differenceMask=Content.");
                state.Require(! HasFlag(item->differenceMask, CompareDirectoriesDiffBit::ContentPending),
                              L"a.bin expected ContentPending not set when sizes differ.");
            }
        }
    }
    else
    {
        state.Require(false, L"Failed to create case folders: content_size_mismatch_no_pending.");
    }

    return state.failure.empty();
});

SelfTest::RunCase(options,
                  suite,
                  L"zero_vs_nonzero_content",
                  [&](SelfTest::CaseState& state) noexcept
{
    // Case: Content compare with zero vs non-zero size does not mark ContentPending.
    if (const auto foldersOpt = CreateCaseFolders(root, L"zero_vs_nonzero_content"))
    {
        const auto& folders = foldersOpt.value();
        state.Require(SelfTest::WriteBinaryFile(folders.left / L"a.bin", {}), L"Failed to create a.bin (left).");
        state.Require(WriteFileFill(folders.right / L"a.bin", 'Z', 1), L"Failed to create a.bin (right).");

        Common::Settings::CompareDirectoriesSettings settings{};
        settings.compareContent = true;

        auto decision = ComputeRootDecision(baseFs, folders, settings, state);
        if (decision)
        {
            const auto* item = FindItem(*decision, L"a.bin");
            state.Require(item != nullptr, L"a.bin missing from decision.");
            if (item)
            {
                state.Require(item->isDifferent, L"a.bin expected isDifferent with compareContent and size mismatch.");
                state.Require(item->selectLeft && item->selectRight, L"a.bin expected select both when compareContent and sizes differ.");
                state.Require(HasFlag(item->differenceMask, CompareDirectoriesDiffBit::Content), L"a.bin expected differenceMask=Content.");
                state.Require(! HasFlag(item->differenceMask, CompareDirectoriesDiffBit::ContentPending),
                              L"a.bin expected ContentPending not set when sizes differ.");
            }
        }
    }
    else
    {
        state.Require(false, L"Failed to create case folders: zero_vs_nonzero_content.");
    }

    return state.failure.empty();
});

SelfTest::RunCase(options,
                  suite,
                  L"unicode_filenames",
                  [&](SelfTest::CaseState& state) noexcept
{
    // Case: Unicode filenames (CJK + emoji) appear in decisions and content compare works.
    if (const auto foldersOpt = CreateCaseFolders(root, L"unicode_filenames"))
    {
        const auto& folders = foldersOpt.value();
        state.Require(SelfTest::WriteTextFile(folders.left / L"\u3053\u3093\u306B\u3061\u306F.txt", "JP"),
                      L"Failed to create ã“ã‚“ã«ã¡ã¯.txt (left).");
        state.Require(SelfTest::WriteTextFile(folders.right / L"\u3053\u3093\u306B\u3061\u306F.txt", "JP"),
                      L"Failed to create ã“ã‚“ã«ã¡ã¯.txt (right).");

        state.Require(WriteFileFill(folders.left / L"emoji_\U0001F600.bin", 'A', 64), L"Failed to create emoji file (left).");
        state.Require(WriteFileFill(folders.right / L"emoji_\U0001F600.bin", 'B', 64), L"Failed to create emoji file (right).");

        Common::Settings::CompareDirectoriesSettings settings{};
        settings.compareContent     = true;
        settings.keepIdenticalItems = true;

        auto session  = std::make_shared<CompareDirectoriesSession>(baseFs, baseFs, folders.left, folders.right, settings);
        auto decision = WaitForContentCompare(session, std::filesystem::path{}, L"emoji_\U0001F600.bin", state);
        if (decision)
        {
            {
                const auto* item = FindItem(*decision, L"\u3053\u3093\u306B\u3061\u306F.txt");
                state.Require(item != nullptr, L"Unicode file missing from decision.");
                if (item)
                {
                    state.Require(! item->isDifferent, L"Unicode identical file expected not different.");
                }
            }
            {
                const auto* item = FindItem(*decision, L"emoji_\U0001F600.bin");
                state.Require(item != nullptr, L"Emoji file missing from decision.");
                if (item)
                {
                    state.Require(item->isDifferent, L"Emoji file expected different with compareContent.");
                    state.Require(item->selectLeft && item->selectRight, L"Emoji file expected select both when content differs.");
                    state.Require(HasFlag(item->differenceMask, CompareDirectoriesDiffBit::Content), L"Emoji file expected differenceMask=Content.");
                    state.Require(! HasFlag(item->differenceMask, CompareDirectoriesDiffBit::ContentPending),
                                  L"Emoji file expected ContentPending cleared after compare completes.");
                }
            }
        }
    }
    else
    {
        state.Require(false, L"Failed to create case folders: unicode_filenames.");
    }

    return state.failure.empty();
});

SelfTest::RunCase(options,
                  suite,
                  L"content short reads",
                  [&](SelfTest::CaseState& state) noexcept
{
    // Case: Content compare tolerates short reads for equal files.
    if (const auto foldersOpt = CreateCaseFolders(root, L"content_shortreads"))
    {
        const auto& folders = foldersOpt.value();
        state.Require(WriteFileFill(folders.left / L"a.bin", 'Z', 4096), L"Failed to create a.bin (left).");
        state.Require(WriteFileFill(folders.right / L"a.bin", 'Z', 4096), L"Failed to create a.bin (right).");

        Common::Settings::CompareDirectoriesSettings settings{};
        settings.compareContent     = true;
        settings.keepIdenticalItems = true;

        wil::com_ptr<IFileSystem> wrapped = CreateShortReadFileSystem(baseFs, folders.left, 1u, 0u);
        state.Require(static_cast<bool>(wrapped), L"Failed to create short-read file system wrapper.");

        const wil::com_ptr<IFileSystem> compareFs = wrapped ? wrapped : baseFs;
        auto session                              = std::make_shared<CompareDirectoriesSession>(compareFs, compareFs, folders.left, folders.right, settings);
        auto decision                             = WaitForContentCompare(session, std::filesystem::path{}, L"a.bin", state);
        if (decision)
        {
            const auto* item = FindItem(*decision, L"a.bin");
            state.Require(item != nullptr, L"a.bin missing from decision.");
            if (item)
            {
                state.Require(! item->isDifferent, L"a.bin expected not different for equal content with short reads.");
                state.Require(! HasFlag(item->differenceMask, CompareDirectoriesDiffBit::Content),
                              L"a.bin expected Content bit cleared for equal content with short reads.");
                state.Require(! HasFlag(item->differenceMask, CompareDirectoriesDiffBit::ContentPending),
                              L"a.bin expected ContentPending cleared after compare completes (short reads).");
                state.Require(! item->selectLeft && ! item->selectRight, L"a.bin expected no selection when equal.");
            }
        }
    }
    else
    {
        state.Require(false, L"Failed to create case folders: content_shortreads.");
    }

    return state.failure.empty();
});

SelfTest::RunCase(options,
                  suite,
                  L"subdir pending",
                  [&](SelfTest::CaseState& state) noexcept
{
    // Case: Subdirectory pending state + flush updates ancestors without navigation.
    if (const auto foldersOpt = CreateCaseFolders(root, L"subdir_pending"))
    {
        const auto& folders = foldersOpt.value();
        state.Require(SelfTest::EnsureDirectory(folders.left / L"sub"), L"Failed to create sub (left).");
        state.Require(SelfTest::EnsureDirectory(folders.right / L"sub"), L"Failed to create sub (right).");
        state.Require(WriteFileFill(folders.left / L"sub" / L"a.bin", 'A', 512 * 1024), L"Failed to create sub\\a.bin (left).");
        state.Require(WriteFileFill(folders.right / L"sub" / L"a.bin", 'A', 512 * 1024), L"Failed to create sub\\a.bin (right).");

        Common::Settings::CompareDirectoriesSettings settings{};
        settings.compareContent        = true;
        settings.compareSubdirectories = true;
        settings.keepIdenticalItems    = true;

        wil::com_ptr<IFileSystem> wrapped = CreateShortReadFileSystem(baseFs, folders.left, 1024u, 1u);
        state.Require(static_cast<bool>(wrapped), L"Failed to create short-read file system wrapper (subdir pending).");

        const wil::com_ptr<IFileSystem> compareFs = wrapped ? wrapped : baseFs;
        auto session                              = std::make_shared<CompareDirectoriesSession>(compareFs, compareFs, folders.left, folders.right, settings);

        std::mutex progressMutex;
        std::condition_variable progressCv;
        bool contentDone = false;

        session->SetContentProgressCallback([&](uint32_t,
                                                const std::filesystem::path&,
                                                std::wstring_view,
                                                uint64_t,
                                                uint64_t,
                                                uint64_t,
                                                uint64_t,
                                                uint64_t pendingContentCompares,
                                                uint64_t totalContentCompares,
                                                uint64_t completedContentCompares) noexcept
        {
            if (pendingContentCompares != 0u || totalContentCompares == 0u || completedContentCompares != totalContentCompares)
            {
                return;
            }

            std::lock_guard lock(progressMutex);
            contentDone = true;
            progressCv.notify_all();
        });

        auto rootDecision = session->GetOrComputeDecision(std::filesystem::path{});
        state.Require(static_cast<bool>(rootDecision), L"subdir pending: root decision is null.");
        if (rootDecision)
        {
            const auto* subItem = FindItem(*rootDecision, L"sub");
            state.Require(subItem != nullptr, L"subdir pending: sub missing from root decision.");
            if (subItem)
            {
                state.Require(subItem->isDirectory, L"subdir pending: sub expected isDirectory.");
                state.Require(HasFlag(subItem->differenceMask, CompareDirectoriesDiffBit::SubdirPending),
                              L"subdir pending: sub expected SubdirPending while content compare is running.");
                state.Require(! HasFlag(subItem->differenceMask, CompareDirectoriesDiffBit::SubdirContent),
                              L"subdir pending: sub expected no SubdirContent while only content compares are pending.");
                state.Require(! subItem->isDifferent, L"subdir pending: sub expected not different while pending.");
                state.Require(! subItem->selectLeft && ! subItem->selectRight, L"subdir pending: sub expected not selected while pending.");
            }
        }

        auto subDecision = session->GetOrComputeDecision(std::filesystem::path(L"sub"));
        state.Require(static_cast<bool>(subDecision), L"subdir pending: sub decision is null.");
        if (subDecision)
        {
            const auto* fileItem = FindItem(*subDecision, L"a.bin");
            state.Require(fileItem != nullptr, L"subdir pending: a.bin missing from sub decision.");
            if (fileItem)
            {
                state.Require(HasFlag(fileItem->differenceMask, CompareDirectoriesDiffBit::ContentPending),
                              L"subdir pending: a.bin expected ContentPending while content compare is running.");
                state.Require(! HasFlag(fileItem->differenceMask, CompareDirectoriesDiffBit::Content),
                              L"subdir pending: a.bin expected no Content bit while pending.");
                state.Require(! fileItem->isDifferent, L"subdir pending: a.bin expected not different while pending.");
                state.Require(! fileItem->selectLeft && ! fileItem->selectRight, L"subdir pending: a.bin expected not selected while pending.");
            }
        }

        {
            std::unique_lock lock(progressMutex);
            static_cast<void>(progressCv.wait_for(lock, std::chrono::milliseconds(SelfTest::ScaleTimeout(30'000)), [&] { return contentDone; }));
        }
        state.Require(contentDone, L"subdir pending: timed out waiting for content compare to finish.");

        // Root decision remains in pending state until pending updates are flushed.
        auto rootBeforeFlush = session->GetOrComputeDecision(std::filesystem::path{});
        if (rootBeforeFlush)
        {
            const auto* subItem = FindItem(*rootBeforeFlush, L"sub");
            if (subItem)
            {
                state.Require(HasFlag(subItem->differenceMask, CompareDirectoriesDiffBit::SubdirPending),
                              L"subdir pending: expected SubdirPending to remain until FlushPendingContentCompareUpdates.");
            }
        }

        session->FlushPendingContentCompareUpdates();
        session->SetContentProgressCallback({});

        auto rootAfterFlush = session->GetOrComputeDecision(std::filesystem::path{});
        state.Require(static_cast<bool>(rootAfterFlush), L"subdir pending: root decision missing after flush.");
        if (rootAfterFlush)
        {
            const auto* subItem = FindItem(*rootAfterFlush, L"sub");
            state.Require(subItem != nullptr, L"subdir pending: sub missing after flush.");
            if (subItem)
            {
                state.Require(subItem->differenceMask == 0u, L"subdir pending: sub expected no difference mask after flush (equal subtree).");
                state.Require(! subItem->isDifferent, L"subdir pending: sub expected not different after flush (equal subtree).");
                state.Require(! subItem->selectLeft && ! subItem->selectRight, L"subdir pending: sub expected not selected after flush (equal subtree).");
            }
        }

        auto subAfterFlush = session->GetOrComputeDecision(std::filesystem::path(L"sub"));
        if (subAfterFlush)
        {
            const auto* fileItem = FindItem(*subAfterFlush, L"a.bin");
            state.Require(fileItem != nullptr, L"subdir pending: a.bin missing after flush.");
            if (fileItem)
            {
                state.Require(fileItem->differenceMask == 0u, L"subdir pending: a.bin expected no difference mask after flush (equal).");
                state.Require(! fileItem->isDifferent, L"subdir pending: a.bin expected not different after flush (equal).");
                state.Require(! fileItem->selectLeft && ! fileItem->selectRight, L"subdir pending: a.bin expected not selected after flush (equal).");
            }
        }
    }
    else
    {
        state.Require(false, L"Failed to create case folders: subdir_pending.");
    }

    return state.failure.empty();
});

SelfTest::RunCase(options,
                  suite,
                  L"subdirs",
                  [&](SelfTest::CaseState& state) noexcept
{
    // Case: Subdirectory content compare selects both directories.
    if (const auto foldersOpt = CreateCaseFolders(root, L"subdirs"))
    {
        const auto& folders = foldersOpt.value();
        state.Require(SelfTest::EnsureDirectory(folders.left / L"sub"), L"Failed to create sub (left).");
        state.Require(SelfTest::EnsureDirectory(folders.right / L"sub"), L"Failed to create sub (right).");
        state.Require(SelfTest::WriteTextFile(folders.left / L"sub" / L"child.txt", "C"), L"Failed to create sub\\child.txt (left).");

        Common::Settings::CompareDirectoriesSettings settings{};
        settings.compareSubdirectories = true;

        auto session = std::make_shared<CompareDirectoriesSession>(baseFs, baseFs, folders.left, folders.right, settings);
        state.Require(static_cast<bool>(session), L"subdirs: failed to create session.");
        state.Require(StartScanAndWaitForIdle(session, std::chrono::milliseconds{SelfTest::ScaleTimeout(10'000)}),
                      L"subdirs: scan did not become idle within timeout.");
        state.Require(DrainPendingSubdirUpdates(session, 256), L"subdirs: failed to drain pending subtree updates.");

        auto decision = session->GetOrComputeDecision(std::filesystem::path{});
        if (decision)
        {
            const auto* item = FindItem(*decision, L"sub");
            state.Require(item != nullptr, L"sub missing from decision.");
            if (item)
            {
                state.Require(item->isDirectory, L"sub expected isDirectory.");
                state.Require(item->isDifferent, L"sub expected isDifferent with compareSubdirectories.");
                state.Require(item->selectLeft && item->selectRight, L"sub expected select both when content differs.");
                state.Require(HasFlag(item->differenceMask, CompareDirectoriesDiffBit::SubdirContent), L"sub expected differenceMask=SubdirContent.");
            }
        }
    }
    else
    {
        state.Require(false, L"Failed to create case folders: subdirs.");
    }

    return state.failure.empty();
});

SelfTest::RunCase(options,
                  suite,
                  L"select_subdirs_only_in_one_pane",
                  [&](SelfTest::CaseState& state) noexcept
{
    // Case: one-sided directory selection follows selectSubdirsOnlyInOnePane.
    if (const auto foldersOpt = CreateCaseFolders(root, L"select_subdirs_only_in_one_pane"))
    {
        const auto& folders = foldersOpt.value();
        state.Require(SelfTest::EnsureDirectory(folders.left / L"left_only_dir"), L"select subdirs: failed to create left_only_dir.");
        state.Require(SelfTest::WriteTextFile(folders.left / L"left_only_dir" / L"child.txt", "L"), L"select subdirs: failed to write child.");

        const auto assertSelection = [&](bool selectSubdirsOnlyInOnePane) noexcept
        {
            Common::Settings::CompareDirectoriesSettings settings{};
            settings.compareSubdirectories      = true;
            settings.selectSubdirsOnlyInOnePane = selectSubdirsOnlyInOnePane;

            auto session = std::make_shared<CompareDirectoriesSession>(baseFs, baseFs, folders.left, folders.right, settings);
            state.Require(StartScanAndWaitForIdle(session, std::chrono::milliseconds{SelfTest::ScaleTimeout(10'000)}),
                          L"select subdirs: scan did not become idle.");
            state.Require(DrainPendingSubdirUpdates(session, 256), L"select subdirs: pending subdir updates did not drain.");

            auto decision = session->GetOrComputeDecision(std::filesystem::path{});
            state.Require(static_cast<bool>(decision), L"select subdirs: decision is null.");
            if (! decision)
            {
                return;
            }

            const auto* item = FindItem(*decision, L"left_only_dir");
            state.Require(item != nullptr, L"select subdirs: left_only_dir missing.");
            if (! item)
            {
                return;
            }

            state.Require(item->isDirectory, L"select subdirs: expected directory item.");
            state.Require(item->isDifferent, L"select subdirs: expected one-sided directory to be different.");
            state.Require(HasFlag(item->differenceMask, CompareDirectoriesDiffBit::OnlyInLeft), L"select subdirs: expected OnlyInLeft bit.");
            state.Require(item->selectLeft == selectSubdirsOnlyInOnePane, L"select subdirs: selectLeft must follow selectSubdirsOnlyInOnePane.");
            state.Require(! item->selectRight, L"select subdirs: right side must not be selected.");
        };

        assertSelection(false);
        assertSelection(true);
    }
    else
    {
        state.Require(false, L"Failed to create case folders: select_subdirs_only_in_one_pane.");
    }

    return state.failure.empty();
});

SelfTest::RunCase(options,
                  suite,
                  L"sync_manifest_nested_differences_only",
                  [&](SelfTest::CaseState& state) noexcept
{
    // Case: Compare sync expands a selected folder into exact per-item destinations, not a recursive whole-folder copy.
    if (const auto foldersOpt = CreateCaseFolders(root, L"sync_manifest_nested_differences_only"))
    {
        const auto& folders = foldersOpt.value();
        state.Require(SelfTest::EnsureDirectory(folders.left / L"sub" / L"left-only-dir"), L"sync manifest: failed to create left-only-dir.");
        state.Require(SelfTest::EnsureDirectory(folders.right / L"sub"), L"sync manifest: failed to create right sub.");
        state.Require(WriteFileFill(folders.left / L"sub" / L"diff.bin", 'L', 8), L"sync manifest: failed to write left diff.bin.");
        state.Require(WriteFileFill(folders.right / L"sub" / L"diff.bin", 'R', 3), L"sync manifest: failed to write right diff.bin.");
        state.Require(SelfTest::WriteTextFile(folders.left / L"sub" / L"same.txt", "same"), L"sync manifest: failed to write left same.txt.");
        state.Require(SelfTest::WriteTextFile(folders.right / L"sub" / L"same.txt", "same"), L"sync manifest: failed to write right same.txt.");
        state.Require(SelfTest::WriteTextFile(folders.left / L"sub" / L"left-only.txt", "left"), L"sync manifest: failed to write left-only.txt.");
        state.Require(SelfTest::WriteTextFile(folders.left / L"sub" / L"left-only-dir" / L"nested.txt", "nested"),
                      L"sync manifest: failed to write nested.txt.");
        state.Require(SelfTest::WriteTextFile(folders.right / L"sub" / L"right-only.txt", "right"), L"sync manifest: failed to write right-only.txt.");

        const std::filesystem::path reparseTarget = folders.left / L"target";
        state.Require(SelfTest::EnsureDirectory(reparseTarget), L"sync manifest: failed to create reparse target.");
        state.Require(SelfTest::WriteTextFile(reparseTarget / L"child.txt", "linked"), L"sync manifest: failed to write reparse target child.");
        const std::filesystem::path leftOnlyReparsePath = folders.left / L"sub" / L"left-only-reparse";
        const bool leftOnlyReparseCreated               = TryCreateDirectorySymlink(leftOnlyReparsePath, reparseTarget);
        if (! leftOnlyReparseCreated)
        {
            const DWORD err = ::GetLastError();
            if (err == ERROR_PRIVILEGE_NOT_HELD || err == ERROR_ACCESS_DENIED || err == ERROR_INVALID_PARAMETER)
            {
                Debug::Warning(L"CompareSelfTest: skipping sync manifest source-only reparse assertion (CreateSymbolicLinkW failed: {}).", err);
            }
            else
            {
                state.Require(false, std::format(L"sync manifest: CreateSymbolicLinkW failed unexpectedly: {}.", err));
            }
        }

        Common::Settings::CompareDirectoriesSettings settings{};
        settings.compareSize           = true;
        settings.compareDateTime       = false;
        settings.compareAttributes     = false;
        settings.compareContent        = false;
        settings.compareSubdirectories = true;

        auto session = std::make_shared<CompareDirectoriesSession>(baseFs, baseFs, folders.left, folders.right, settings);
        state.Require(StartScanAndWaitForIdle(session, std::chrono::milliseconds{SelfTest::ScaleTimeout(10'000)}), L"sync manifest: scan did not become idle.");
        state.Require(DrainPendingSubdirUpdates(session, 256), L"sync manifest: failed to drain pending subtree updates.");
        state.Require(static_cast<bool>(session->GetOrComputeDecision(std::filesystem::path(L"sub"))), L"sync manifest: sub decision is null.");

        CompareSyncManifest manifest{};
        CompareSyncManifestBlocker blocker{};
        const CompareSyncManifestStatus status = session->TryBuildSyncManifest(ComparePane::Left, {std::filesystem::path(L"sub")}, manifest, blocker);
        state.Require(
            status == CompareSyncManifestStatus::Ready,
            std::format(L"sync manifest: expected Ready, got {} hr=0x{:08X}.", static_cast<unsigned int>(status), static_cast<unsigned long>(blocker.hr)));

        const auto findItem = [&](CompareSyncManifestItemKind kind, const std::filesystem::path& relativePath) noexcept -> const CompareSyncManifestItem*
        {
            const std::wstring expected = relativePath.lexically_normal().generic_wstring();
            for (const CompareSyncManifestItem& item : manifest.items)
            {
                if (item.kind == kind &&
                    CompareStringOrdinal(item.relativePath.lexically_normal().generic_wstring().c_str(), -1, expected.c_str(), -1, TRUE) == CSTR_EQUAL)
                {
                    return &item;
                }
            }
            return nullptr;
        };

        const CompareSyncManifestItem* shell = findItem(CompareSyncManifestItemKind::DirectoryShell, L"sub");
        state.Require(shell != nullptr, L"sync manifest: expected directory shell for sub.");
        const CompareSyncManifestItem* diff = findItem(CompareSyncManifestItemKind::File, L"sub/diff.bin");
        state.Require(diff != nullptr, L"sync manifest: expected diff.bin file transfer.");
        const CompareSyncManifestItem* leftOnly = findItem(CompareSyncManifestItemKind::File, L"sub/left-only.txt");
        state.Require(leftOnly != nullptr, L"sync manifest: expected left-only.txt file transfer.");
        const CompareSyncManifestItem* leftOnlyDir = findItem(CompareSyncManifestItemKind::WholeSubtree, L"sub/left-only-dir");
        state.Require(leftOnlyDir != nullptr, L"sync manifest: expected whole subtree transfer for left-only-dir.");
        const CompareSyncManifestItem* leftOnlyReparse = findItem(CompareSyncManifestItemKind::File, L"sub/left-only-reparse");
        if (leftOnlyReparseCreated)
        {
            state.Require(leftOnlyReparse != nullptr, L"sync manifest: source-only directory reparse must transfer as a normal item.");
            state.Require(findItem(CompareSyncManifestItemKind::WholeSubtree, L"sub/left-only-reparse") == nullptr,
                          L"sync manifest: source-only directory reparse must not be emitted as a recursive whole subtree.");
        }
        state.Require(findItem(CompareSyncManifestItemKind::File, L"sub/same.txt") == nullptr, L"sync manifest: identical same.txt must not be transferred.");
        state.Require(findItem(CompareSyncManifestItemKind::File, L"sub/right-only.txt") == nullptr,
                      L"sync manifest: destination-only right-only.txt must not be transferred from left.");

        if (diff)
        {
            state.Require(diff->sourceAbsolutePath == folders.left / L"sub" / L"diff.bin", L"sync manifest: diff source path mismatch.");
            state.Require(diff->destinationAbsolutePath == folders.right / L"sub" / L"diff.bin", L"sync manifest: diff destination path mismatch.");
            state.Require((diff->flags & FILESYSTEM_FLAG_RECURSIVE) == 0, L"sync manifest: file item must not be recursive.");
        }
        if (leftOnlyDir)
        {
            state.Require(leftOnlyDir->sourceAbsolutePath == folders.left / L"sub" / L"left-only-dir", L"sync manifest: left-only-dir source path mismatch.");
            state.Require(leftOnlyDir->destinationAbsolutePath == folders.right / L"sub" / L"left-only-dir",
                          L"sync manifest: left-only-dir destination path mismatch.");
            state.Require((leftOnlyDir->flags & FILESYSTEM_FLAG_RECURSIVE) != 0, L"sync manifest: whole subtree item must be recursive.");
        }
        if (leftOnlyReparse)
        {
            state.Require((leftOnlyReparse->flags & FILESYSTEM_FLAG_RECURSIVE) == 0, L"sync manifest: source-only directory reparse must not be recursive.");
        }
    }
    else
    {
        state.Require(false, L"Failed to create case folders: sync_manifest_nested_differences_only.");
    }

    return state.failure.empty();
});

SelfTest::RunCase(options,
                  suite,
                  L"sync_manifest_move_preserves_identical_children",
                  [&](SelfTest::CaseState& state) noexcept
{
    // Case: a selected existing directory expands to changed children only, so Move does not delete identical source children.
    if (const auto foldersOpt = CreateCaseFolders(root, L"sync_manifest_move_preserves_identical_children"))
    {
        const auto& folders = foldersOpt.value();
        state.Require(SelfTest::EnsureDirectory(folders.left / L"move"), L"sync move manifest: failed to create left move.");
        state.Require(SelfTest::EnsureDirectory(folders.right / L"move"), L"sync move manifest: failed to create right move.");
        state.Require(WriteFileFill(folders.left / L"move" / L"changed.bin", 'L', 9), L"sync move manifest: failed to write left changed.bin.");
        state.Require(WriteFileFill(folders.right / L"move" / L"changed.bin", 'R', 2), L"sync move manifest: failed to write right changed.bin.");
        state.Require(SelfTest::WriteTextFile(folders.left / L"move" / L"identical.txt", "same"), L"sync move manifest: failed to write left identical.txt.");
        state.Require(SelfTest::WriteTextFile(folders.right / L"move" / L"identical.txt", "same"), L"sync move manifest: failed to write right identical.txt.");

        Common::Settings::CompareDirectoriesSettings settings{};
        settings.compareSize           = true;
        settings.compareDateTime       = false;
        settings.compareAttributes     = false;
        settings.compareContent        = false;
        settings.compareSubdirectories = true;

        auto session = std::make_shared<CompareDirectoriesSession>(baseFs, baseFs, folders.left, folders.right, settings);
        state.Require(StartScanAndWaitForIdle(session, std::chrono::milliseconds{SelfTest::ScaleTimeout(10'000)}),
                      L"sync move manifest: scan did not become idle.");
        state.Require(DrainPendingSubdirUpdates(session, 256), L"sync move manifest: failed to drain pending subtree updates.");
        state.Require(static_cast<bool>(session->GetOrComputeDecision(std::filesystem::path(L"move"))), L"sync move manifest: move decision is null.");

        CompareSyncManifest manifest{};
        CompareSyncManifestBlocker blocker{};
        const CompareSyncManifestStatus status = session->TryBuildSyncManifest(ComparePane::Left, {std::filesystem::path(L"move")}, manifest, blocker);
        state.Require(
            status == CompareSyncManifestStatus::Ready,
            std::format(L"sync move manifest: expected Ready, got {} hr=0x{:08X}.", static_cast<unsigned int>(status), static_cast<unsigned long>(blocker.hr)));

        const auto contains = [&](CompareSyncManifestItemKind kind, const std::filesystem::path& relativePath) noexcept
        {
            const std::wstring expected = relativePath.lexically_normal().generic_wstring();
            return std::any_of(manifest.items.begin(),
                               manifest.items.end(),
                               [&](const CompareSyncManifestItem& item) noexcept
            {
                return item.kind == kind &&
                       CompareStringOrdinal(item.relativePath.lexically_normal().generic_wstring().c_str(), -1, expected.c_str(), -1, TRUE) == CSTR_EQUAL;
            });
        };

        state.Require(contains(CompareSyncManifestItemKind::DirectoryShell, L"move"), L"sync move manifest: expected move directory shell.");
        state.Require(contains(CompareSyncManifestItemKind::File, L"move/changed.bin"), L"sync move manifest: expected changed.bin file transfer.");
        state.Require(! contains(CompareSyncManifestItemKind::WholeSubtree, L"move"),
                      L"sync move manifest: existing move directory must not be whole-subtree.");
        state.Require(! contains(CompareSyncManifestItemKind::File, L"move/identical.txt"),
                      L"sync move manifest: identical child must not be included in move manifest.");
    }
    else
    {
        state.Require(false, L"Failed to create case folders: sync_manifest_move_preserves_identical_children.");
    }

    return state.failure.empty();
});

SelfTest::RunCase(options,
                  suite,
                  L"sync_manifest_not_ready_without_cached_decision",
                  [&](SelfTest::CaseState& state) noexcept
{
    // Case: sync planning is cache-only; missing decisions block instead of producing a partial manifest.
    if (const auto foldersOpt = CreateCaseFolders(root, L"sync_manifest_not_ready_without_cached_decision"))
    {
        const auto& folders = foldersOpt.value();
        state.Require(SelfTest::EnsureDirectory(folders.left / L"sub"), L"sync manifest not-ready: failed to create left sub.");
        state.Require(SelfTest::EnsureDirectory(folders.right / L"sub"), L"sync manifest not-ready: failed to create right sub.");
        state.Require(SelfTest::WriteTextFile(folders.left / L"sub" / L"changed.txt", "left"), L"sync manifest not-ready: failed to create left changed.txt.");
        state.Require(SelfTest::WriteTextFile(folders.right / L"sub" / L"changed.txt", "right"),
                      L"sync manifest not-ready: failed to create right changed.txt.");

        Common::Settings::CompareDirectoriesSettings settings{};
        settings.compareSize           = true;
        settings.compareDateTime       = false;
        settings.compareAttributes     = false;
        settings.compareContent        = false;
        settings.compareSubdirectories = true;

        auto session = std::make_shared<CompareDirectoriesSession>(baseFs, baseFs, folders.left, folders.right, settings);

        CompareSyncManifest manifest{};
        CompareSyncManifestBlocker blocker{};
        const CompareSyncManifestStatus status = session->TryBuildSyncManifest(ComparePane::Left, {std::filesystem::path(L"sub")}, manifest, blocker);
        state.Require(status == CompareSyncManifestStatus::NotReady,
                      std::format(L"sync manifest not-ready: expected NotReady, got {} hr=0x{:08X}.",
                                  static_cast<unsigned int>(status),
                                  static_cast<unsigned long>(blocker.hr)));
        state.Require(blocker.status == CompareSyncManifestStatus::NotReady, L"sync manifest not-ready: blocker status mismatch.");
        state.Require(blocker.relativePath.empty(), L"sync manifest not-ready: expected missing root decision blocker.");
        state.Require(blocker.hr == S_FALSE, L"sync manifest not-ready: expected S_FALSE blocker hr.");
        state.Require(manifest.items.empty(), L"sync manifest not-ready: partial manifest items must not be returned.");
    }
    else
    {
        state.Require(false, L"Failed to create case folders: sync_manifest_not_ready_without_cached_decision.");
    }

    return state.failure.empty();
});

SelfTest::RunCase(options,
                  suite,
                  L"sync_manifest_pending_blocks_and_schedules",
                  [&](SelfTest::CaseState& state) noexcept
{
    // Case: pending content/subdir decisions block sync and map to high-priority scan requests.
    if (const auto foldersOpt = CreateCaseFolders(root, L"sync_manifest_pending_blocks_and_schedules"))
    {
        const auto& folders = foldersOpt.value();
        state.Require(SelfTest::EnsureDirectory(folders.left / L"sub"), L"sync manifest pending: failed to create left sub.");
        state.Require(SelfTest::EnsureDirectory(folders.right / L"sub"), L"sync manifest pending: failed to create right sub.");
        state.Require(WriteFileFill(folders.left / L"sub" / L"a.bin", 'A', 512 * 1024), L"sync manifest pending: failed to create left a.bin.");
        state.Require(WriteFileFill(folders.right / L"sub" / L"a.bin", 'A', 512 * 1024), L"sync manifest pending: failed to create right a.bin.");

        Common::Settings::CompareDirectoriesSettings settings{};
        settings.compareContent        = true;
        settings.compareSubdirectories = true;
        settings.keepIdenticalItems    = true;

        wil::com_ptr<IFileSystem> wrapped = CreateShortReadFileSystem(baseFs, folders.left, 4096u, static_cast<DWORD>(SelfTest::ScaleTimeout(2)));
        state.Require(static_cast<bool>(wrapped), L"sync manifest pending: failed to create short-read wrapper.");

        auto session =
            std::make_shared<CompareDirectoriesSession>(wrapped ? wrapped : baseFs, wrapped ? wrapped : baseFs, folders.left, folders.right, settings);

        std::mutex progressMutex;
        std::condition_variable progressCv;
        bool contentDone = false;

        session->SetContentProgressCallback([&](uint32_t,
                                                const std::filesystem::path&,
                                                std::wstring_view,
                                                uint64_t,
                                                uint64_t,
                                                uint64_t,
                                                uint64_t,
                                                uint64_t pendingContentCompares,
                                                uint64_t totalContentCompares,
                                                uint64_t completedContentCompares) noexcept
        {
            if (pendingContentCompares != 0u || totalContentCompares == 0u || completedContentCompares != totalContentCompares)
            {
                return;
            }

            std::lock_guard lock(progressMutex);
            contentDone = true;
            progressCv.notify_all();
        });

        auto rootDecision = session->GetOrComputeDecision(std::filesystem::path{});
        state.Require(static_cast<bool>(rootDecision), L"sync manifest pending: root decision is null.");
        if (rootDecision)
        {
            const auto* subItem = FindItem(*rootDecision, L"sub");
            state.Require(subItem != nullptr, L"sync manifest pending: sub missing from root decision.");
            if (subItem)
            {
                state.Require(HasFlag(subItem->differenceMask, CompareDirectoriesDiffBit::SubdirPending),
                              L"sync manifest pending: sub expected SubdirPending.");
            }
        }

        auto subDecision = session->GetOrComputeDecision(std::filesystem::path(L"sub"));
        state.Require(static_cast<bool>(subDecision), L"sync manifest pending: sub decision is null.");
        if (subDecision)
        {
            const auto* fileItem = FindItem(*subDecision, L"a.bin");
            state.Require(fileItem != nullptr, L"sync manifest pending: a.bin missing from sub decision.");
            if (fileItem)
            {
                state.Require(HasFlag(fileItem->differenceMask, CompareDirectoriesDiffBit::ContentPending),
                              L"sync manifest pending: a.bin expected ContentPending.");
            }
        }

        CompareSyncManifest subManifest{};
        CompareSyncManifestBlocker subBlocker{};
        const CompareSyncManifestStatus subStatus = session->TryBuildSyncManifest(ComparePane::Left, {std::filesystem::path(L"sub")}, subManifest, subBlocker);
        state.Require(subStatus == CompareSyncManifestStatus::NotReady,
                      std::format(L"sync manifest pending: expected SubdirPending NotReady, got status={} reason={}.",
                                  static_cast<unsigned int>(subStatus),
                                  static_cast<unsigned int>(subBlocker.reason)));
        state.Require(subBlocker.reason == CompareSyncManifestBlockerReason::SubdirPending,
                      L"sync manifest pending: selected folder should block on SubdirPending.");
        state.Require(subBlocker.relativePath == std::filesystem::path(L"sub"), L"sync manifest pending: SubdirPending blocker path mismatch.");
        state.Require(subManifest.items.empty(), L"sync manifest pending: SubdirPending blocker returned partial items.");

        const std::filesystem::path fileRelative = std::filesystem::path(L"sub") / L"a.bin";
        CompareSyncManifest fileManifest{};
        CompareSyncManifestBlocker fileBlocker{};
        const CompareSyncManifestStatus fileStatus = session->TryBuildSyncManifest(ComparePane::Left, {fileRelative}, fileManifest, fileBlocker);
        state.Require(fileStatus == CompareSyncManifestStatus::NotReady,
                      std::format(L"sync manifest pending: expected ContentPending NotReady, got status={} reason={}.",
                                  static_cast<unsigned int>(fileStatus),
                                  static_cast<unsigned int>(fileBlocker.reason)));
        state.Require(fileBlocker.reason == CompareSyncManifestBlockerReason::ContentPending,
                      L"sync manifest pending: selected file should block on ContentPending.");
        state.Require(fileBlocker.relativePath == fileRelative, L"sync manifest pending: ContentPending blocker path mismatch.");
        state.Require(fileManifest.items.empty(), L"sync manifest pending: ContentPending blocker returned partial items.");

        session->StartScan();
        const CompareDirectoriesPerfStats beforeSchedule = session->GetPerfStats();
        const auto scheduleLikeWindow                    = [&](const CompareSyncManifestBlocker& blocker) noexcept
        {
            session->RequestScanForFolder(blocker.relativePath);
            const std::filesystem::path parent = blocker.relativePath.parent_path().lexically_normal();
            if (parent != blocker.relativePath)
            {
                session->RequestScanForFolder(parent);
            }
        };
        scheduleLikeWindow(subBlocker);
        scheduleLikeWindow(fileBlocker);
        const CompareDirectoriesPerfStats afterSchedule = session->GetPerfStats();
        state.Require(afterSchedule.scanQueueHighHighWater > beforeSchedule.scanQueueHighHighWater ||
                          afterSchedule.scanScheduledHighWater > beforeSchedule.scanScheduledHighWater,
                      L"sync manifest pending: blocker scheduling did not enqueue high-priority scan work.");

        {
            std::unique_lock lock(progressMutex);
            static_cast<void>(progressCv.wait_for(lock, std::chrono::milliseconds(SelfTest::ScaleTimeout(30'000)), [&] { return contentDone; }));
        }
        state.Require(contentDone, L"sync manifest pending: timed out waiting for content compare cleanup.");
        session->FlushPendingContentCompareUpdates();
        session->SetContentProgressCallback({});
    }
    else
    {
        state.Require(false, L"Failed to create case folders: sync_manifest_pending_blocks_and_schedules.");
    }

    return state.failure.empty();
});

SelfTest::RunCase(options,
                  suite,
                  L"no_sync_deep_scan",
                  [&](SelfTest::CaseState& state) noexcept
{
    // Case: GetOrComputeDecision() must not perform a synchronous deep subtree traversal.
    if (const auto foldersOpt = CreateCaseFolders(root, L"no_sync_deep_scan"))
    {
        const auto& folders = foldersOpt.value();

        const std::filesystem::path leftRoot  = folders.left;
        const std::filesystem::path rightRoot = folders.right;

        state.Require(SelfTest::EnsureDirectory(leftRoot / L"sub" / L"sub2"), L"no_sync_deep_scan: failed to create sub tree (left).");
        state.Require(SelfTest::EnsureDirectory(rightRoot / L"sub" / L"sub2"), L"no_sync_deep_scan: failed to create sub tree (right).");
        state.Require(SelfTest::WriteTextFile(leftRoot / L"sub" / L"sub2" / L"leaf.txt", "L"), L"no_sync_deep_scan: failed to create leaf.txt (left).");

        std::atomic_uint32_t readDirCalls{0};

        const wil::com_ptr<IFileSystem> leftFs  = CreateCountingReadDirectoryFileSystem(baseFs, &readDirCalls);
        const wil::com_ptr<IFileSystem> rightFs = CreateCountingReadDirectoryFileSystem(baseFs, &readDirCalls);
        state.Require(static_cast<bool>(leftFs), L"no_sync_deep_scan: failed to create left counting fs.");
        state.Require(static_cast<bool>(rightFs), L"no_sync_deep_scan: failed to create right counting fs.");

        Common::Settings::CompareDirectoriesSettings settings{};
        settings.compareSubdirectories = true;

        const uint32_t before = readDirCalls.load(std::memory_order_acquire);
        auto session          = std::make_shared<CompareDirectoriesSession>(leftFs, rightFs, leftRoot, rightRoot, settings);
        auto decision         = session->GetOrComputeDecision(std::filesystem::path{});
        state.Require(static_cast<bool>(decision), L"no_sync_deep_scan: root decision is null.");
        const uint32_t after = readDirCalls.load(std::memory_order_acquire);

        state.Require((after - before) == 2u,
                      std::format(L"no_sync_deep_scan: expected exactly 2 ReadDirectoryInfo calls (root left+right), got {}.", (after - before)));
    }
    else
    {
        state.Require(false, L"Failed to create case folders: no_sync_deep_scan.");
    }

    return state.failure.empty();
});

SelfTest::RunCase(options,
                  suite,
                  L"subdirattrs",
                  [&](SelfTest::CaseState& state) noexcept
{
    // Case: Compare attributes of subdirectories selects both.
    if (const auto foldersOpt = CreateCaseFolders(root, L"subdirattrs"))
    {
        const auto& folders = foldersOpt.value();
        state.Require(SelfTest::EnsureDirectory(folders.left / L"sub"), L"Failed to create sub (left).");
        state.Require(SelfTest::EnsureDirectory(folders.right / L"sub"), L"Failed to create sub (right).");

        const std::filesystem::path leftDir = folders.left / L"sub";
        const DWORD leftAttrs               = ::GetFileAttributesW(leftDir.c_str());
        state.Require(leftAttrs != INVALID_FILE_ATTRIBUTES, L"GetFileAttributesW failed for sub (left).");
        if (leftAttrs != INVALID_FILE_ATTRIBUTES)
        {
            state.Require(::SetFileAttributesW(leftDir.c_str(), leftAttrs | FILE_ATTRIBUTE_HIDDEN) != 0, L"SetFileAttributesW failed for sub (left).");
        }

        Common::Settings::CompareDirectoriesSettings settings{};
        settings.compareSubdirectoryAttributes = true;

        auto decision = ComputeRootDecision(baseFs, folders, settings, state);
        if (decision)
        {
            const auto* item = FindItem(*decision, L"sub");
            state.Require(item != nullptr, L"sub missing from decision.");
            if (item)
            {
                state.Require(item->isDirectory, L"sub expected isDirectory.");
                state.Require(item->isDifferent, L"sub expected isDifferent with compareSubdirectoryAttributes.");
                state.Require(item->selectLeft && item->selectRight, L"sub expected select both when attributes differ.");
                state.Require(HasFlag(item->differenceMask, CompareDirectoriesDiffBit::SubdirAttributes), L"sub expected differenceMask=SubdirAttributes.");
            }
        }
    }
    else
    {
        state.Require(false, L"Failed to create case folders: subdirattrs.");
    }

    return state.failure.empty();
});

SelfTest::RunCase(options,
                  suite,
                  L"missing folder",
                  [&](SelfTest::CaseState& state) noexcept
{
    // Case: Missing folder is reported without failing the decision.
    if (const auto foldersOpt = CreateCaseFolders(root, L"missing_folder"))
    {
        const auto& folders = foldersOpt.value();
        state.Require(SelfTest::EnsureDirectory(folders.left / L"sub"), L"Failed to create sub (left).");
        state.Require(SelfTest::WriteTextFile(folders.left / L"sub" / L"a.txt", "A"), L"Failed to create sub\\a.txt (left).");

        Common::Settings::CompareDirectoriesSettings settings{};
        auto session  = std::make_shared<CompareDirectoriesSession>(baseFs, baseFs, folders.left, folders.right, settings);
        auto decision = session->GetOrComputeDecision(std::filesystem::path(L"sub"));
        state.Require(static_cast<bool>(decision), L"missing folder: decision is null.");
        if (decision)
        {
            state.Require(SUCCEEDED(decision->hr), L"missing folder: expected decision hr success.");
            state.Require(! decision->leftFolderMissing, L"missing folder: expected leftFolderMissing=false.");
            state.Require(decision->rightFolderMissing, L"missing folder: expected rightFolderMissing=true.");
        }
    }
    else
    {
        state.Require(false, L"Failed to create case folders: missing_folder.");
    }

    return state.failure.empty();
});

SelfTest::RunCase(options,
                  suite,
                  L"missing_side_empty_enumeration",
                  [&](SelfTest::CaseState& state) noexcept
{
    // Case: the compare wrapper enumerates a missing-side folder as an empty success.
    if (const auto foldersOpt = CreateCaseFolders(root, L"missing_side_empty_enumeration"))
    {
        const auto& folders = foldersOpt.value();
        state.Require(SelfTest::EnsureDirectory(folders.left / L"sub"), L"missing side empty: failed to create left sub.");
        state.Require(SelfTest::WriteTextFile(folders.left / L"sub" / L"a.txt", "A"), L"missing side empty: failed to write left sub file.");

        auto session = std::make_shared<CompareDirectoriesSession>(baseFs, baseFs, folders.left, folders.right, Common::Settings::CompareDirectoriesSettings{});
        session->SetCompareEnabled(true);
        const auto fsRight = CreateCompareDirectoriesFileSystem(ComparePane::Right, session);
        state.Require(static_cast<bool>(fsRight), L"missing side empty: failed to create right compare wrapper.");

        wil::com_ptr<IFilesInformation> info;
        const HRESULT hr = fsRight->ReadDirectoryInfo((folders.right / L"sub").c_str(), info.put());
        state.Require(SUCCEEDED(hr), std::format(L"missing side empty: ReadDirectoryInfo failed. hr=0x{:08X}", static_cast<unsigned long>(hr)));
        state.Require(static_cast<bool>(info), L"missing side empty: info is null.");
        if (info)
        {
            unsigned long count   = 1;
            const HRESULT countHr = info->GetCount(&count);
            state.Require(SUCCEEDED(countHr), L"missing side empty: GetCount failed.");
            state.Require(count == 0u, L"missing side empty: missing-side enumeration must be empty.");
        }
    }
    else
    {
        state.Require(false, L"Failed to create case folders: missing_side_empty_enumeration.");
    }

    return state.failure.empty();
});

SelfTest::RunCase(options,
                  suite,
                  L"failed_enumeration_retries_without_version_bump",
                  [&](SelfTest::CaseState& state) noexcept
{
    // Case: Transient enumeration failures are returned to the caller but not cached.
    if (const auto foldersOpt = CreateCaseFolders(root, L"failed_enumeration_retries_without_version_bump"))
    {
        const auto& folders = foldersOpt.value();
        state.Require(SelfTest::WriteTextFile(folders.left / L"retry.txt", "L"), L"Failed to create retry.txt (left).");

        auto forcedHr = std::make_shared<std::atomic<HRESULT>>(E_ACCESSDENIED);

        ReadDirectoryTestBehavior behavior{};
        behavior.targetPath       = folders.left;
        behavior.forcedHrOverride = forcedHr;

        wil::com_ptr<IFileSystem> flakyLeft = CreateReadDirectoryBehaviorFileSystem(baseFs, behavior);
        state.Require(static_cast<bool>(flakyLeft), L"Failed to create retry ReadDirectoryInfo wrapper.");

        auto session = std::make_shared<CompareDirectoriesSession>(
            flakyLeft ? flakyLeft : baseFs, baseFs, folders.left, folders.right, Common::Settings::CompareDirectoriesSettings{});
        const uint64_t versionBefore = session->GetVersion();

        const auto failedDecision = session->GetOrComputeDecision(std::filesystem::path{});
        state.Require(static_cast<bool>(failedDecision), L"failed_enumeration_retries_without_version_bump: failed decision is null.");
        if (failedDecision)
        {
            state.Require(failedDecision->hr == E_ACCESSDENIED,
                          std::format(L"failed_enumeration_retries_without_version_bump: expected E_ACCESSDENIED, got 0x{:08X}.",
                                      static_cast<unsigned long>(failedDecision->hr)));
        }

        const CompareDirectoriesPerfStats failedStats = session->GetPerfStats();
        state.Require(failedStats.decisionCacheEntries == 0u, L"failed_enumeration_retries_without_version_bump: failed decision should not be cached.");
        state.Require(failedStats.decisionCacheEstimatedBytes == 0u,
                      L"failed_enumeration_retries_without_version_bump: failed decision should not contribute cache bytes.");

        forcedHr->store(S_OK, std::memory_order_release);

        const auto successDecision = session->GetOrComputeDecision(std::filesystem::path{});
        state.Require(static_cast<bool>(successDecision), L"failed_enumeration_retries_without_version_bump: retry decision is null.");
        state.Require(session->GetVersion() == versionBefore, L"failed_enumeration_retries_without_version_bump: retry should not need a version bump.");
        if (successDecision)
        {
            state.Require(SUCCEEDED(successDecision->hr), L"failed_enumeration_retries_without_version_bump: retry should succeed.");
            const auto* item = FindItem(*successDecision, L"retry.txt");
            state.Require(item != nullptr, L"failed_enumeration_retries_without_version_bump: retry.txt missing after retry.");
            if (item)
            {
                state.Require(item->existsLeft && ! item->existsRight, L"failed_enumeration_retries_without_version_bump: retry.txt should be left-only.");
                state.Require(HasFlag(item->differenceMask, CompareDirectoriesDiffBit::OnlyInLeft),
                              L"failed_enumeration_retries_without_version_bump: retry.txt expected OnlyInLeft.");
            }
        }

        const CompareDirectoriesPerfStats successStats = session->GetPerfStats();
        state.Require(successStats.decisionCacheEntries != 0u, L"failed_enumeration_retries_without_version_bump: successful retry should be cacheable.");
    }
    else
    {
        state.Require(false, L"Failed to create case folders: failed_enumeration_retries_without_version_bump.");
    }

    return state.failure.empty();
});

SelfTest::RunCase(options,
                  suite,
                  L"reparse",
                  [&](SelfTest::CaseState& state) noexcept
{
    // Case: Reparse points are not traversed for subdirectory comparison.
    if (const auto foldersOpt = CreateCaseFolders(root, L"reparse"))
    {
        const auto& folders                = foldersOpt.value();
        const std::filesystem::path target = folders.left / L"target";
        state.Require(SelfTest::EnsureDirectory(target), L"Failed to create reparse target (left).");
        state.Require(SelfTest::WriteTextFile(target / L"child.txt", "C"), L"Failed to create target\\child.txt (left).");

        const std::filesystem::path linkPath = folders.left / L"sub";
        const bool linkCreated               = TryCreateDirectorySymlink(linkPath, target);
        if (! linkCreated)
        {
            const DWORD err = ::GetLastError();
            if (err == ERROR_PRIVILEGE_NOT_HELD || err == ERROR_ACCESS_DENIED || err == ERROR_INVALID_PARAMETER)
            {
                Debug::Warning(L"CompareSelfTest: skipping reparse point test (CreateSymbolicLinkW failed: {0}).", err);
            }
            else
            {
                state.Require(false, std::format(L"CreateSymbolicLinkW failed unexpectedly: {}.", err));
            }
        }
        else
        {
            state.Require(SelfTest::EnsureDirectory(folders.right / L"sub"), L"Failed to create sub directory (right).");

            Common::Settings::CompareDirectoriesSettings settings{};
            settings.compareSubdirectories = true;
            settings.keepIdenticalItems    = true;

            auto decision = ComputeRootDecision(baseFs, folders, settings, state);
            if (decision)
            {
                const auto* item = FindItem(*decision, L"sub");
                state.Require(item != nullptr, L"sub missing from decision.");
                if (item)
                {
                    state.Require(item->isDirectory, L"sub expected isDirectory.");
                    state.Require(! HasFlag(item->differenceMask, CompareDirectoriesDiffBit::SubdirContent),
                                  L"sub expected SubdirContent not set for reparse points.");
                }
            }
        }
    }
    else
    {
        state.Require(false, L"Failed to create case folders: reparse.");
    }

    return state.failure.empty();
});

SelfTest::RunCase(options,
                  suite,
                  L"dummy_content",
                  [&](SelfTest::CaseState& state) noexcept
{
    // Case: Dummy filesystem paths use plugin I/O for content compare (cross-filesystem support).
    if (dummyFs && dummyIo && dummyOps)
    {
        const SelfTest::TestSandbox sandbox = SelfTest::AcquireTestSandbox(SelfTest::SelfTestSuite::CompareDirectories, L"dummy_content");
        if (! state.Require(sandbox.IsValid(), L"Dummy: failed to acquire TestSandbox root (content compare)."))
        {
            return state.failure.empty();
        }

        const std::filesystem::path baseRoot  = sandbox.root;
        const std::filesystem::path leftRoot  = baseRoot / L"left";
        const std::filesystem::path rightRoot = baseRoot / L"right";
        state.Require(EnsureDirectoryExistsFsOps(dummyOps, leftRoot), L"Dummy: failed to create left root.");
        state.Require(EnsureDirectoryExistsFsOps(dummyOps, rightRoot), L"Dummy: failed to create right root.");

        state.Require(WriteFileTextFsIo(dummyIo, leftRoot / L"a.bin", "SAME"), L"Dummy: failed to write a.bin (left).");
        state.Require(WriteFileTextFsIo(dummyIo, rightRoot / L"a.bin", "SAME"), L"Dummy: failed to write a.bin (right).");

        Common::Settings::CompareDirectoriesSettings settings{};
        settings.compareContent     = true;
        settings.keepIdenticalItems = true;

        auto session  = std::make_shared<CompareDirectoriesSession>(dummyFs, dummyFs, leftRoot, rightRoot, settings);
        auto decision = WaitForContentCompare(session, std::filesystem::path{}, L"a.bin", state);
        if (decision)
        {
            const auto* item = FindItem(*decision, L"a.bin");
            state.Require(item != nullptr, L"Dummy: a.bin missing from decision.");
            if (item)
            {
                state.Require(! item->isDifferent, L"Dummy: a.bin expected identical after content compare.");
                state.Require(item->differenceMask == 0u, L"Dummy: a.bin expected differenceMask=0 after content compare.");
            }
        }
    }
    else
    {
        state.Require(false, L"CompareSelfTest: FileSystemDummy unavailable for cross-filesystem content compare test.");
    }

    return state.failure.empty();
});

SelfTest::RunCase(options,
                  suite,
                  L"normalized_name_collision_preserves_same_side_entries",
                  [&](SelfTest::CaseState& state) noexcept
{
    // Case: Same-side entries that collide after Win32 trailing dot/space normalization must not overwrite each other.
    if (dummyFs && dummyIo && dummyOps)
    {
        const SelfTest::TestSandbox sandbox =
            SelfTest::AcquireTestSandbox(SelfTest::SelfTestSuite::CompareDirectories, L"normalized_name_collision_preserves_same_side_entries");
        if (! state.Require(sandbox.IsValid(), L"Normalized collision: failed to acquire TestSandbox root."))
        {
            return state.failure.empty();
        }

        const std::filesystem::path baseRoot  = sandbox.root;
        const std::filesystem::path leftRoot  = baseRoot / L"left";
        const std::filesystem::path rightRoot = baseRoot / L"right";
        state.Require(EnsureDirectoryExistsFsOps(dummyOps, leftRoot), L"Normalized collision: failed to create left root.");
        state.Require(EnsureDirectoryExistsFsOps(dummyOps, rightRoot), L"Normalized collision: failed to create right root.");

        state.Require(WriteFileTextFsIo(dummyIo, leftRoot / L"report.", "DOT"), L"Normalized collision: failed to write report. (left).");
        state.Require(WriteFileTextFsIo(dummyIo, leftRoot / L"report", "SAME"), L"Normalized collision: failed to write report (left).");
        state.Require(WriteFileTextFsIo(dummyIo, rightRoot / L"report", "SAME"), L"Normalized collision: failed to write report (right).");
        state.Require(WriteFileTextFsIo(dummyIo, baseRoot / L"sentinel.txt", "OUT"), L"Normalized collision: failed to write out-of-scope sentinel.");

        Common::Settings::CompareDirectoriesSettings settings{};
        settings.keepIdenticalItems = true;

        auto session = std::make_shared<CompareDirectoriesSession>(dummyFs, dummyFs, leftRoot, rightRoot, settings);
        std::shared_ptr<const CompareDirectoriesFolderDecision> decision;
        if (! TryGetRootDecisionWithSeh(*session, decision))
        {
            state.Require(false, L"Normalized collision: GetOrComputeDecision crashed.");
        }

        state.Require(static_cast<bool>(decision), L"Normalized collision: decision is null.");
        if (decision)
        {
            state.Require(SUCCEEDED(decision->hr), L"Normalized collision: decision hr is failure.");
            state.Require(FindItem(*decision, L"sentinel.txt") == nullptr, L"Normalized collision: out-of-scope sentinel leaked into decision.");

            const auto* pairedItem = FindItem(*decision, L"report");
            state.Require(pairedItem != nullptr, L"Normalized collision: report missing from decision.");
            if (pairedItem)
            {
                state.Require(pairedItem->existsLeft && pairedItem->existsRight, L"Normalized collision: report expected on both sides.");
                state.Require(! pairedItem->isDifferent, L"Normalized collision: report should pair the untrimmed same-name entries.");
                state.Require(pairedItem->leftSizeBytes == 4u && pairedItem->rightSizeBytes == 4u,
                              L"Normalized collision: report sizes should come from the untrimmed entries.");
            }

            const auto* dotItem = FindItem(*decision, L"report.");
            state.Require(dotItem != nullptr, L"Normalized collision: report. missing from decision.");
            if (dotItem)
            {
                state.Require(dotItem->existsLeft && ! dotItem->existsRight, L"Normalized collision: report. expected only on the left side.");
                state.Require(dotItem->isDifferent, L"Normalized collision: report. expected isDifferent.");
                state.Require(HasFlag(dotItem->differenceMask, CompareDirectoriesDiffBit::OnlyInLeft),
                              L"Normalized collision: report. expected differenceMask=OnlyInLeft.");
                state.Require(dotItem->leftSizeBytes == 3u, L"Normalized collision: report. size should come from the trailing-dot entry.");
            }
        }
    }
    else
    {
        state.Require(false, L"CompareSelfTest: FileSystemDummy unavailable for normalized name collision test.");
    }

    return state.failure.empty();
});

SelfTest::RunCase(options,
                  suite,
                  L"deep_tree",
                  [&](SelfTest::CaseState& state) noexcept
{
    // Case: Deep directory trees do not overflow the stack (iterative traversal).
    if (dummyFs && dummyIo && dummyOps)
    {
        const SelfTest::TestSandbox sandbox = SelfTest::AcquireTestSandbox(SelfTest::SelfTestSuite::CompareDirectories, L"deep_tree");
        if (! state.Require(sandbox.IsValid(), L"Dummy: failed to acquire TestSandbox root (deep_tree)."))
        {
            return state.failure.empty();
        }

        const std::filesystem::path baseRoot  = sandbox.root;
        const std::filesystem::path leftRoot  = baseRoot / L"left";
        const std::filesystem::path rightRoot = baseRoot / L"right";
        state.Require(EnsureDirectoryExistsFsOps(dummyOps, leftRoot), L"Dummy: failed to create deep left root.");
        state.Require(EnsureDirectoryExistsFsOps(dummyOps, rightRoot), L"Dummy: failed to create deep right root.");

        constexpr size_t kDepth = 1024;

        std::filesystem::path leftPath  = leftRoot;
        std::filesystem::path rightPath = rightRoot;
        for (size_t i = 0; i < kDepth; ++i)
        {
            const std::wstring name = std::format(L"d{:04}", i);
            leftPath /= name;
            rightPath /= name;
            const HRESULT leftHr  = dummyOps->CreateDirectory(leftPath.c_str());
            const HRESULT rightHr = dummyOps->CreateDirectory(rightPath.c_str());
            state.Require(SUCCEEDED(leftHr) || leftHr == HRESULT_FROM_WIN32(ERROR_ALREADY_EXISTS),
                          std::format(L"Dummy: failed to create left dir at depth {}.", i));
            state.Require(SUCCEEDED(rightHr) || rightHr == HRESULT_FROM_WIN32(ERROR_ALREADY_EXISTS),
                          std::format(L"Dummy: failed to create right dir at depth {}.", i));
        }

        state.Require(WriteFileTextFsIo(dummyIo, leftPath / L"leaf.txt", "L"), L"Dummy: failed to create leaf.txt (left).");

        Common::Settings::CompareDirectoriesSettings settings{};
        settings.compareSubdirectories = true;

        auto session = std::make_shared<CompareDirectoriesSession>(dummyFs, dummyFs, leftRoot, rightRoot, settings);
        state.Require(static_cast<bool>(session), L"Dummy: failed to create session (deep_tree).");
        state.Require(StartScanAndWaitForIdle(session, std::chrono::milliseconds{SelfTest::ScaleTimeout(60'000)}),
                      L"Dummy: scan did not become idle within timeout (deep_tree).");
        state.Require(DrainPendingSubdirUpdates(session, 512), L"Dummy: failed to drain pending subtree updates (deep_tree).");

        const auto decision = session->GetOrComputeDecision(std::filesystem::path{});
        if (decision)
        {
            const auto* item = FindItem(*decision, L"d0000");
            state.Require(item != nullptr, L"Dummy: d0000 missing from decision.");
            if (item)
            {
                state.Require(item->isDirectory, L"Dummy: d0000 expected isDirectory.");
                state.Require(item->isDifferent, L"Dummy: d0000 expected isDifferent from deep leaf mismatch.");
                state.Require(HasFlag(item->differenceMask, CompareDirectoriesDiffBit::SubdirContent),
                              L"Dummy: d0000 expected differenceMask=SubdirContent from deep leaf mismatch.");
            }
        }
    }
    else
    {
        state.Require(false, L"CompareSelfTest: FileSystemDummy unavailable for deep tree test.");
    }

    return state.failure.empty();
});

SelfTest::RunCase(options,
                  suite,
                  L"invalidate",
                  [&](SelfTest::CaseState& state) noexcept
{
    // Case: Version invalidation mid-scan does not cache stale results.
    if (dummyFs && dummyIo && dummyOps)
    {
        const SelfTest::TestSandbox sandbox = SelfTest::AcquireTestSandbox(SelfTest::SelfTestSuite::CompareDirectories, L"invalidate");
        if (! state.Require(sandbox.IsValid(), L"Dummy: failed to acquire TestSandbox root (invalidate)."))
        {
            return state.failure.empty();
        }

        const std::filesystem::path baseRoot  = sandbox.root;
        const std::filesystem::path leftRoot  = baseRoot / L"left";
        const std::filesystem::path rightRoot = baseRoot / L"right";
        state.Require(EnsureDirectoryExistsFsOps(dummyOps, leftRoot), L"Dummy: failed to create invalidate left root.");
        state.Require(EnsureDirectoryExistsFsOps(dummyOps, rightRoot), L"Dummy: failed to create invalidate right root.");

        constexpr size_t kDepth         = 256;
        std::filesystem::path leftPath  = leftRoot;
        std::filesystem::path rightPath = rightRoot;
        for (size_t i = 0; i < kDepth; ++i)
        {
            const std::wstring name = std::format(L"d{}", i);
            leftPath /= name;
            rightPath /= name;
            static_cast<void>(dummyOps->CreateDirectory(leftPath.c_str()));
            static_cast<void>(dummyOps->CreateDirectory(rightPath.c_str()));
        }
        state.Require(WriteFileTextFsIo(dummyIo, leftPath / L"leaf.txt", "X"), L"Dummy: failed to create invalidate leaf.txt (left).");

        Common::Settings::CompareDirectoriesSettings settings{};
        settings.compareSubdirectories = true;

        auto session                 = std::make_shared<CompareDirectoriesSession>(dummyFs, dummyFs, leftRoot, rightRoot, settings);
        const uint64_t versionBefore = session->GetVersion();

        const auto decisionBefore = session->GetOrComputeDecision(std::filesystem::path{});
        state.Require(static_cast<bool>(decisionBefore), L"Invalidate: initial decision missing.");

        std::mutex progressMutex;
        std::condition_variable progressCv;
        bool scanInProgress = false;

        session->SetScanProgressCallback(
            [&](const std::filesystem::path&, std::wstring_view, uint64_t scannedFolders, uint64_t, uint32_t activeScans, uint64_t, uint64_t) noexcept
        {
            if (scannedFolders == 0u || activeScans == 0u)
            {
                return;
            }

            std::lock_guard lock(progressMutex);
            scanInProgress = true;
            progressCv.notify_all();
        });

        session->StartScan();

        {
            std::unique_lock lock(progressMutex);
            static_cast<void>(progressCv.wait_for(lock, std::chrono::milliseconds{SelfTest::ScaleTimeout(5'000)}, [&] { return scanInProgress; }));
        }
        session->SetScanProgressCallback({});

        state.Require(scanInProgress, L"Invalidate: scan did not start within timeout.");

        session->Invalidate();
        state.Require(session->GetVersion() == versionBefore + 1u, L"Invalidate: expected version bump.");

        const auto decisionAfter = session->GetOrComputeDecision(std::filesystem::path{});
        state.Require(static_cast<bool>(decisionAfter), L"Invalidate: decision missing after invalidation.");
        if (decisionBefore && decisionAfter)
        {
            state.Require(decisionAfter != decisionBefore, L"Invalidate: expected a new decision after invalidation (stale result cached).");
        }

        state.Require(StartScanAndWaitForIdle(session, std::chrono::milliseconds{SelfTest::ScaleTimeout(30'000)}),
                      L"Invalidate: scan did not become idle within timeout after invalidation.");

        const CompareDirectoriesPerfStats stats = session->GetPerfStats();
        state.Require(stats.scanActiveScans == 0u, L"Invalidate: expected scanActiveScans == 0 after idle.");
        state.Require(stats.scanQueueSize == 0u, L"Invalidate: expected scanQueueSize == 0 after idle.");
        state.Require(stats.scanScheduledKeys == 0u, L"Invalidate: expected scanScheduledKeys == 0 after idle.");
        state.Require(stats.scanInFlightKeys == 0u, L"Invalidate: expected scanInFlightKeys == 0 after idle.");
    }
    else
    {
        state.Require(false, L"CompareSelfTest: FileSystemDummy unavailable for invalidation test.");
    }

    return state.failure.empty();
});

SelfTest::RunCase(options,
                  suite,
                  L"concurrent_get_or_compute_decision",
                  [&](SelfTest::CaseState& state) noexcept
{
    // Case: Concurrent GetOrComputeDecision and Invalidate does not crash and never returns null.
    if (const auto foldersOpt = CreateCaseFolders(root, L"concurrent_get_or_compute_decision"))
    {
        const auto& folders = foldersOpt.value();
        state.Require(SelfTest::WriteTextFile(folders.left / L"a.txt", "A"), L"Failed to create a.txt (left).");
        state.Require(SelfTest::WriteTextFile(folders.right / L"a.txt", "A"), L"Failed to create a.txt (right).");
        state.Require(SelfTest::WriteTextFile(folders.left / L"b.txt", "L"), L"Failed to create b.txt (left).");
        state.Require(SelfTest::WriteTextFile(folders.right / L"b.txt", "R"), L"Failed to create b.txt (right).");

        Common::Settings::CompareDirectoriesSettings settings{};
        const auto session = std::make_shared<CompareDirectoriesSession>(baseFs, baseFs, folders.left, folders.right, settings);

        std::atomic<uint32_t> nullDecisions{0};
        constexpr int kWorkerCount     = 4;
        constexpr int kWorkerIters     = 50;
        constexpr int kInvalidateIters = 10;

        std::vector<std::jthread> workers;
        workers.reserve(kWorkerCount);
        for (int i = 0; i < kWorkerCount; ++i)
        {
            workers.emplace_back([session, &nullDecisions](std::stop_token) noexcept
            {
                for (int j = 0; j < kWorkerIters; ++j)
                {
                    auto decision = session->GetOrComputeDecision(std::filesystem::path{});
                    if (! decision)
                    {
                        nullDecisions.fetch_add(1u, std::memory_order_relaxed);
                    }
                }
            });
        }

        std::jthread invalidator([session](std::stop_token) noexcept
        {
            for (int j = 0; j < kInvalidateIters; ++j)
            {
                session->Invalidate();
            }
        });

        for (auto& worker : workers)
        {
            worker.join();
        }
        invalidator.join();

        state.Require(nullDecisions.load(std::memory_order_relaxed) == 0u, L"Concurrent GetOrComputeDecision returned null.");
    }
    else
    {
        state.Require(false, L"Failed to create case folders: concurrent_get_or_compute_decision.");
    }

    return state.failure.empty();
});

SelfTest::RunCase(options,
                  suite,
                  L"empty_directories",
                  [&](SelfTest::CaseState& state) noexcept
{
    // Case: Empty trees produce an empty decision and scan reaches idle.
    if (const auto foldersOpt = CreateCaseFolders(root, L"empty_directories"))
    {
        const auto& folders = foldersOpt.value();

        Common::Settings::CompareDirectoriesSettings settings{};
        const auto session = std::make_shared<CompareDirectoriesSession>(baseFs, baseFs, folders.left, folders.right, settings);

        session->StartScan();
        state.Require(StartScanAndWaitForIdle(session, std::chrono::milliseconds{SelfTest::ScaleTimeout(10'000)}),
                      L"Empty directories: scan did not become idle within timeout.");

        const auto decision = session->GetOrComputeDecision(std::filesystem::path{});
        state.Require(static_cast<bool>(decision), L"Empty directories: decision is null.");
        if (decision)
        {
            state.Require(SUCCEEDED(decision->hr), L"Empty directories: expected decision hr success.");
            state.Require(decision->items.empty(), L"Empty directories: expected no items.");
        }
    }
    else
    {
        state.Require(false, L"Failed to create case folders: empty_directories.");
    }

    return state.failure.empty();
});

SelfTest::RunCase(options,
                  suite,
                  L"ignore",
                  [&](SelfTest::CaseState& state) noexcept
{
    // Case: Ignore patterns exclude files/directories.
    if (const auto foldersOpt = CreateCaseFolders(root, L"ignore"))
    {
        const auto& folders = foldersOpt.value();
        state.Require(SelfTest::WriteTextFile(folders.left / L"ignore.log", "I"), L"Failed to create ignore.log (left).");
        state.Require(SelfTest::WriteTextFile(folders.left / L"keep.txt", "K"), L"Failed to create keep.txt (left).");
        state.Require(SelfTest::EnsureDirectory(folders.left / L"ignore_dir"), L"Failed to create ignore_dir (left).");
        state.Require(SelfTest::EnsureDirectory(folders.left / L"keep_dir"), L"Failed to create keep_dir (left).");

        Common::Settings::CompareDirectoriesSettings settings{};
        settings.ignoreFiles               = true;
        settings.ignoreFilesPatterns       = L"*.log";
        settings.ignoreDirectories         = true;
        settings.ignoreDirectoriesPatterns = L"ignore*";

        auto decision = ComputeRootDecision(baseFs, folders, settings, state);
        if (decision)
        {
            state.Require(FindItem(*decision, L"keep.txt") != nullptr, L"keep.txt expected in decision.");
            state.Require(FindItem(*decision, L"ignore.log") == nullptr, L"ignore.log expected to be ignored.");
            state.Require(FindItem(*decision, L"keep_dir") != nullptr, L"keep_dir expected in decision.");
            state.Require(FindItem(*decision, L"ignore_dir") == nullptr, L"ignore_dir expected to be ignored.");
        }
    }
    else
    {
        state.Require(false, L"Failed to create case folders: ignore.");
    }

    return state.failure.empty();
});

SelfTest::RunCase(options,
                  suite,
                  L"ignore_direct_navigation_subtree",
                  [&](SelfTest::CaseState& state) noexcept
{
    // Case: Direct navigation into an ignored directory subtree must stay excluded.
    if (const auto foldersOpt = CreateCaseFolders(root, L"ignore_direct_navigation_subtree"))
    {
        const auto& folders = foldersOpt.value();
        state.Require(SelfTest::EnsureDirectory(folders.left / L"ignored_dir" / L"nested"), L"Ignore direct: failed to create ignored nested dir (left).");
        state.Require(SelfTest::EnsureDirectory(folders.right / L"ignored_dir" / L"nested"), L"Ignore direct: failed to create ignored nested dir (right).");
        state.Require(SelfTest::WriteTextFile(folders.left / L"ignored_dir" / L"left_only.txt", "L"), L"Ignore direct: failed to write left_only.txt.");
        state.Require(SelfTest::WriteTextFile(folders.right / L"ignored_dir" / L"right_only.txt", "R"), L"Ignore direct: failed to write right_only.txt.");
        state.Require(SelfTest::WriteTextFile(folders.left / L"ignored_dir" / L"nested" / L"left_nested.txt", "L"),
                      L"Ignore direct: failed to write left nested file.");
        state.Require(SelfTest::WriteTextFile(folders.right / L"ignored_dir" / L"nested" / L"right_nested.txt", "R"),
                      L"Ignore direct: failed to write right nested file.");
        state.Require(SelfTest::WriteTextFile(folders.left / L"keep.txt", "L"), L"Ignore direct: failed to write keep.txt (left).");

        Common::Settings::CompareDirectoriesSettings settings{};
        settings.ignoreDirectories         = true;
        settings.ignoreDirectoriesPatterns = L"ignored*";

        auto session = std::make_shared<CompareDirectoriesSession>(baseFs, baseFs, folders.left, folders.right, settings);
        std::shared_ptr<const CompareDirectoriesFolderDecision> rootDecision;
        state.Require(TryGetRootDecisionWithSeh(*session, rootDecision), L"Ignore direct: root GetOrComputeDecision crashed.");
        if (rootDecision)
        {
            state.Require(SUCCEEDED(rootDecision->hr), L"Ignore direct: root decision hr is failure.");
            state.Require(FindItem(*rootDecision, L"keep.txt") != nullptr, L"Ignore direct: keep.txt expected in root decision.");
            state.Require(FindItem(*rootDecision, L"ignored_dir") == nullptr, L"Ignore direct: ignored_dir expected excluded from root decision.");
        }

        const auto ignoredDecision = session->GetOrComputeDecision(std::filesystem::path(L"ignored_dir"));
        state.Require(static_cast<bool>(ignoredDecision), L"Ignore direct: ignored_dir decision is null.");
        if (ignoredDecision)
        {
            state.Require(SUCCEEDED(ignoredDecision->hr), L"Ignore direct: ignored_dir decision hr is failure.");
            state.Require(ignoredDecision->items.empty(), L"Ignore direct: ignored_dir direct decision should be empty.");
        }

        const auto nestedDecision = session->GetOrComputeDecision(std::filesystem::path(L"ignored_dir") / L"nested");
        state.Require(static_cast<bool>(nestedDecision), L"Ignore direct: ignored nested decision is null.");
        if (nestedDecision)
        {
            state.Require(SUCCEEDED(nestedDecision->hr), L"Ignore direct: ignored nested decision hr is failure.");
            state.Require(nestedDecision->items.empty(), L"Ignore direct: ignored nested direct decision should be empty.");
        }
    }
    else
    {
        state.Require(false, L"Failed to create case folders: ignore_direct_navigation_subtree.");
    }

    return state.failure.empty();
});

SelfTest::RunCase(options,
                  suite,
                  L"ignore_multiple_patterns",
                  [&](SelfTest::CaseState& state) noexcept
{
    // Case: Multiple ignore patterns exclude all matching files.
    if (const auto foldersOpt = CreateCaseFolders(root, L"ignore_multiple_patterns"))
    {
        const auto& folders = foldersOpt.value();
        state.Require(SelfTest::WriteTextFile(folders.left / L"ignore.log", "I"), L"Failed to create ignore.log (left).");
        state.Require(SelfTest::WriteTextFile(folders.left / L"foo.tmp", "T"), L"Failed to create foo.tmp (left).");
        state.Require(SelfTest::WriteTextFile(folders.left / L"keep.txt", "K"), L"Failed to create keep.txt (left).");

        Common::Settings::CompareDirectoriesSettings settings{};
        settings.ignoreFiles         = true;
        settings.ignoreFilesPatterns = L"*.log;*.tmp";

        auto decision = ComputeRootDecision(baseFs, folders, settings, state);
        if (decision)
        {
            state.Require(FindItem(*decision, L"keep.txt") != nullptr, L"keep.txt expected in decision.");
            state.Require(FindItem(*decision, L"ignore.log") == nullptr, L"ignore.log expected to be ignored.");
            state.Require(FindItem(*decision, L"foo.tmp") == nullptr, L"foo.tmp expected to be ignored.");
        }
    }
    else
    {
        state.Require(false, L"Failed to create case folders: ignore_multiple_patterns.");
    }

    return state.failure.empty();
});

SelfTest::RunCase(options,
                  suite,
                  L"ignore_pattern_length_cap",
                  [&](SelfTest::CaseState& state) noexcept
{
    // Case: Overly long ignore patterns are dropped (harden against pathological inputs).
    if (const auto foldersOpt = CreateCaseFolders(root, L"ignore_pattern_length_cap"))
    {
        const auto& folders = foldersOpt.value();
        state.Require(SelfTest::WriteTextFile(folders.left / L"cap.txt", "C"), L"Failed to create cap.txt (left).");

        Common::Settings::CompareDirectoriesSettings settings{};
        settings.ignoreFiles         = true;
        settings.ignoreFilesPatterns = std::wstring(129, L'*'); // exceeds cap (128) => dropped

        auto decision = ComputeRootDecision(baseFs, folders, settings, state);
        if (decision)
        {
            state.Require(FindItem(*decision, L"cap.txt") != nullptr, L"cap.txt expected in decision (pattern too long must not ignore).");
        }
    }
    else
    {
        state.Require(false, L"Failed to create case folders: ignore_pattern_length_cap.");
    }

    return state.failure.empty();
});

SelfTest::RunCase(options,
                  suite,
                  L"ignore_pattern_count_cap",
                  [&](SelfTest::CaseState& state) noexcept
{
    // Case: Ignore patterns are capped to a bounded count to keep matching work bounded.
    if (const auto foldersOpt = CreateCaseFolders(root, L"ignore_pattern_count_cap"))
    {
        const auto& folders = foldersOpt.value();
        state.Require(SelfTest::WriteTextFile(folders.left / L"cap.log", "C"), L"Failed to create cap.log (left).");

        std::wstring patterns;
        for (int i = 0; i < 32; ++i)
        {
            patterns.append(std::format(L"p{:02};", i)); // non-matching
        }
        patterns.append(L"*.log"); // would match, but is beyond the cap

        Common::Settings::CompareDirectoriesSettings settings{};
        settings.ignoreFiles         = true;
        settings.ignoreFilesPatterns = patterns;

        auto decision = ComputeRootDecision(baseFs, folders, settings, state);
        if (decision)
        {
            state.Require(FindItem(*decision, L"cap.log") != nullptr, L"cap.log expected in decision (pattern beyond cap must not ignore).");
        }
    }
    else
    {
        state.Require(false, L"Failed to create case folders: ignore_pattern_count_cap.");
    }

    return state.failure.empty();
});

SelfTest::RunCase(options,
                  suite,
                  L"ignore_wildcard_pathology_runtime_bound",
                  [&](SelfTest::CaseState& state) noexcept
{
    // Case: A pathological wildcard pattern list must not hang the scan (work is bounded by pattern caps + matcher budgets).
    if (const auto foldersOpt = CreateCaseFolders(root, L"ignore_wildcard_pathology_runtime_bound"))
    {
        const auto& folders = foldersOpt.value();

        constexpr int kFileCount = 256;
        for (int i = 0; i < kFileCount; ++i)
        {
            const std::wstring name = std::format(L"file{:04}.txt", i);
            state.Require(SelfTest::WriteTextFile(folders.left / name, "X"), L"Failed to create file (left).");
            if (i != 0)
            {
                state.Require(SelfTest::WriteTextFile(folders.right / name, "X"), L"Failed to create file (right).");
            }
        }

        std::wstring longWildcard;
        longWildcard.reserve(128);
        for (int i = 0; i < 64; ++i)
        {
            longWildcard.append(L"*a"); // length 128; requires many 'a' (won't match our file names)
        }

        std::wstring patterns;
        patterns.reserve((longWildcard.size() + 1) * 32);
        for (int i = 0; i < 32; ++i)
        {
            patterns.append(longWildcard);
            patterns.push_back(L';');
        }

        Common::Settings::CompareDirectoriesSettings settings{};
        settings.ignoreFiles         = true;
        settings.ignoreFilesPatterns = patterns;

        const auto session = std::make_shared<CompareDirectoriesSession>(baseFs, baseFs, folders.left, folders.right, settings);
        state.Require(StartScanAndWaitForIdle(session, std::chrono::milliseconds{SelfTest::ScaleTimeout(10'000)}),
                      L"Wildcard patterns: scan did not become idle within timeout.");

        const auto decision = session->GetOrComputeDecision(std::filesystem::path{});
        state.Require(static_cast<bool>(decision), L"Wildcard patterns: decision is null.");
        if (decision)
        {
            state.Require(SUCCEEDED(decision->hr), L"Wildcard patterns: expected decision hr success.");
            const auto* item = FindItem(*decision, L"file0000.txt");
            state.Require(item != nullptr, L"file0000.txt expected in decision.");
            if (item)
            {
                state.Require(item->isDifferent, L"file0000.txt should remain surfaced as a differing item.");
                state.Require(item->existsLeft && ! item->existsRight, L"file0000.txt should be present only on the left side.");
            }
        }
    }
    else
    {
        state.Require(false, L"Failed to create case folders: ignore_wildcard_pathology_runtime_bound.");
    }

    return state.failure.empty();
});

SelfTest::RunCase(options,
                  suite,
                  L"crash_quarantine_synthetic_marker",
                  [&](SelfTest::CaseState& state) noexcept
{
    // Case: A synthetic crash marker should trigger crash-quarantine to disable the last-active filesystem plugin.
    constexpr std::wstring_view kPluginId = L"builtin/file-system-s3";

    const std::filesystem::path sessionPath = SessionState::GetSessionStatePath();
    state.Require(! sessionPath.empty(), L"SessionState::GetSessionStatePath returned empty path.");
    const std::filesystem::path baseDir = sessionPath.parent_path();
    state.Require(! baseDir.empty(), L"SessionState base directory is empty.");

    const std::filesystem::path crashDir   = baseDir / L"Crashes";
    const std::filesystem::path markerPath = crashDir / L"last_crash.txt";

    struct Backup
    {
        bool hadFile = false;
        std::vector<std::byte> bytes;
    };

    const auto readAllBytes = [&](const std::filesystem::path& path, Backup& backup) noexcept -> bool
    {
        backup.hadFile = false;
        backup.bytes.clear();

        std::error_code ec;
        if (! std::filesystem::exists(path, ec))
        {
            return true;
        }

        backup.hadFile = true;

        wil::unique_handle file(CreateFileW(
            path.c_str(), GENERIC_READ, FILE_SHARE_READ | FILE_SHARE_WRITE | FILE_SHARE_DELETE, nullptr, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, nullptr));
        if (! file)
        {
            state.Require(false, L"Failed to open backup file for read.");
            return false;
        }

        LARGE_INTEGER size{};
        if (GetFileSizeEx(file.get(), &size) == 0)
        {
            state.Require(false, L"Failed to get backup file size.");
            return false;
        }

        if (size.QuadPart < 0 || size.QuadPart > static_cast<LONGLONG>((std::numeric_limits<DWORD>::max)()))
        {
            state.Require(false, L"Backup file too large.");
            return false;
        }

        const DWORD bytesToRead = static_cast<DWORD>(size.QuadPart);
        backup.bytes.resize(bytesToRead);

        DWORD readBytes = 0;
        if (bytesToRead > 0)
        {
            if (ReadFile(file.get(), backup.bytes.data(), bytesToRead, &readBytes, nullptr) == 0 || readBytes != bytesToRead)
            {
                state.Require(false, L"Failed to read backup file bytes.");
                return false;
            }
        }

        return true;
    };

    Backup sessionBackup{};
    Backup markerBackup{};
    if (! readAllBytes(sessionPath, sessionBackup) || ! readAllBytes(markerPath, markerBackup))
    {
        return state.failure.empty();
    }

    const auto restore = wil::scope_exit([&] noexcept
    {
        const auto restoreFile = [&](const std::filesystem::path& path, const Backup& backup) noexcept
        {
            std::error_code ec;
            if (! backup.hadFile)
            {
                static_cast<void>(std::filesystem::remove(path, ec));
                return;
            }

            static_cast<void>(std::filesystem::create_directories(path.parent_path(), ec));
            wil::unique_handle file(CreateFileW(path.c_str(), GENERIC_WRITE, FILE_SHARE_READ, nullptr, CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, nullptr));
            if (! file)
            {
                return;
            }

            if (! backup.bytes.empty())
            {
                DWORD written = 0;
                static_cast<void>(WriteFile(file.get(), backup.bytes.data(), static_cast<DWORD>(backup.bytes.size()), &written, nullptr));
            }

            static_cast<void>(FlushFileBuffers(file.get()));
        };

        restoreFile(sessionPath, sessionBackup);
        restoreFile(markerPath, markerBackup);
    });

    const auto writeUtf16File = [&](const std::filesystem::path& path, std::wstring_view content) noexcept -> bool
    {
        std::error_code ec;
        std::filesystem::create_directories(path.parent_path(), ec);
        if (ec)
        {
            state.Require(false, L"Failed to create directory for marker file.");
            return false;
        }

        wil::unique_handle file(CreateFileW(path.c_str(), GENERIC_WRITE, FILE_SHARE_READ, nullptr, CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, nullptr));
        if (! file)
        {
            state.Require(false, L"Failed to create marker file.");
            return false;
        }

        const wchar_t bom = 0xFEFF;
        DWORD written     = 0;
        if (! WriteFile(file.get(), &bom, sizeof(bom), &written, nullptr))
        {
            state.Require(false, L"Failed to write marker BOM.");
            return false;
        }

        const DWORD bytes = static_cast<DWORD>(content.size() * sizeof(wchar_t));
        if (bytes > 0)
        {
            if (! WriteFile(file.get(), content.data(), bytes, &written, nullptr))
            {
                state.Require(false, L"Failed to write marker content.");
                return false;
            }
        }

        static_cast<void>(FlushFileBuffers(file.get()));
        return true;
    };

    std::wstring sessionText;
    sessionText.reserve(32 + kPluginId.size());
    sessionText.append(L"fsPlugin=");
    sessionText.append(kPluginId);
    sessionText.append(L"\r\nop=browse\r\n");

    if (! writeUtf16File(sessionPath, sessionText) || ! writeUtf16File(markerPath, L"synthetic"))
    {
        return state.failure.empty();
    }

    Common::Settings::Settings settings{};
    settings.plugins.currentFileSystemPluginId = std::wstring(kPluginId);

    const bool oldAutoAccept = HostGetAutoAcceptPrompts();
    HostSetAutoAcceptPrompts(true);
    const auto restoreAutoAccept = wil::scope_exit([&] noexcept { HostSetAutoAcceptPrompts(oldAutoAccept); });

    CrashQuarantine::OfferPluginDisableIfPreviousCrashDetected(settings);

    const auto isDisabled = [&](std::wstring_view id) noexcept -> bool
    {
        for (const std::wstring& disabled : settings.plugins.disabledPluginIds)
        {
            if (OrdinalString::EqualsNoCase(disabled, id))
            {
                return true;
            }
        }
        return false;
    };

    state.Require(isDisabled(kPluginId), L"Crash quarantine did not disable the expected plugin id.");
    state.Require(settings.plugins.currentFileSystemPluginId.empty(), L"Crash quarantine did not clear currentFileSystemPluginId.");

    return state.failure.empty();
});
