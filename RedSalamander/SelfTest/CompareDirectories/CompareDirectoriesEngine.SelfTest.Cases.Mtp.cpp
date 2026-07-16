const auto AcquireMtpJournalLocalAppDataSandbox = [&](SelfTest::CaseState& state,
                                                      std::wstring_view caseName,
                                                      std::wstring_view label,
                                                      std::wstring& localAppData,
                                                      std::wstring& previousLocalAppData) noexcept -> bool
{
    previousLocalAppData = GetEnvVarTrimmed(L"LOCALAPPDATA");

    const SelfTest::TestSandbox sandbox = SelfTest::AcquireTestSandbox(SelfTest::SelfTestSuite::CompareDirectories, caseName);
    state.Require(sandbox.IsValid(), std::format(L"{}: failed to acquire TestSandbox scratch root for MTP journal state.", label));
    if (! sandbox.IsValid())
    {
        return false;
    }

    localAppData = sandbox.root.wstring();
    if (! SetEnvironmentVariableW(L"LOCALAPPDATA", localAppData.c_str()))
    {
        state.Require(false,
                      std::format(L"{}: failed to redirect LOCALAPPDATA to TestSandbox. error={}", label, static_cast<unsigned long>(GetLastError())));
        return false;
    }

    AppendCompareSelfTestTraceLine(std::format(L"{}: redirected MTP journal LOCALAPPDATA to '{}'", label, localAppData));
    return true;
};

const auto RestoreMtpJournalLocalAppDataSandbox = [](const std::wstring& previousLocalAppData) noexcept
{
    static_cast<void>(SetEnvironmentVariableW(L"LOCALAPPDATA", previousLocalAppData.empty() ? nullptr : previousLocalAppData.c_str()));
};

const auto extractJsonUInt      = SelfTest::ExtractJsonUInt;
const auto narrowAscii          = SelfTest::NarrowAscii;
const auto stableDeviceHash     = SelfTest::StableDeviceHash;
const auto ensureDirectoryExists = SelfTest::EnsureDirectoryExists;
const auto writeUtf8File        = SelfTest::WriteUtf8File;
const auto notifyInjectedJournal = [](std::wstring_view deviceIdentity) noexcept
{
    using NotifyInjectedJournalFunc = HRESULT(__stdcall*)(const wchar_t* deviceIdentity);
    const std::wstring identity(deviceIdentity);
    return SelfTest::CallMtpPluginExport<NotifyInjectedJournalFunc>(
        "RedSalamanderMtpNotifyJournalInjectedForSelfTest", identity.c_str());
};

SelfTest::RunCase(options,
                  suite,
                  L"mtp_factory_single_mode_id_contract",
                  [&](SelfTest::CaseState& state) noexcept
{
    const FileSystemPluginManager::PluginEntry* entry = FindFileSystemPluginById(kBuiltinMtpFileSystemId);
    state.Require(entry != nullptr && ! entry->path.empty(), L"MTP factory: plugin entry not found.");
    if (entry == nullptr || entry->path.empty())
    {
        return false;
    }

    wil::unique_hmodule module(LoadLibraryExW(entry->path.c_str(), nullptr, LOAD_WITH_ALTERED_SEARCH_PATH));
    state.Require(static_cast<bool>(module), std::format(L"MTP factory: LoadLibraryExW failed. error={}", static_cast<unsigned long>(GetLastError())));
    if (! module)
    {
        return false;
    }

#pragma warning(push)
#pragma warning(disable : 4191) // unsafe conversion from FARPROC
    const auto createFactory = reinterpret_cast<CreateFactoryFunc>(GetProcAddress(module.get(), "RedSalamanderCreate"));
    const auto getSchema     = reinterpret_cast<GetConfigurationSchemaFunc>(GetProcAddress(module.get(), "RedSalamanderGetConfigurationSchema"));
#pragma warning(pop)
    state.Require(createFactory != nullptr, L"MTP factory: RedSalamanderCreate export not found.");
    state.Require(getSchema != nullptr, L"MTP factory: RedSalamanderGetConfigurationSchema export not found.");
    if (createFactory == nullptr || getSchema == nullptr)
    {
        return false;
    }

    FactoryOptions factoryOptions{};
    factoryOptions.debugLevel = DEBUG_LEVEL_NONE;

    const auto requireCreateAccepted = [&](const wchar_t* pluginId, std::wstring_view label) noexcept
    {
        wil::com_ptr<IFileSystem> fs;
        const HRESULT hr = createFactory(__uuidof(IFileSystem), &factoryOptions, GetHostServices(), pluginId, fs.put_void());
        state.Require(SUCCEEDED(hr) && fs, std::format(L"MTP factory: {} create expected success, got hr=0x{:08X}.", label, static_cast<unsigned long>(hr)));
    };

    requireCreateAccepted(nullptr, L"nullptr id");
    requireCreateAccepted(L"", L"empty id");
    requireCreateAccepted(L"builtin/file-system-mtp", L"canonical id");
    requireCreateAccepted(L"BUILTIN/FILE-SYSTEM-MTP", L"case-insensitive id");

    wil::com_ptr<IFileSystem> invalidIdFs;
    const HRESULT invalidIdHr =
        createFactory(__uuidof(IFileSystem), &factoryOptions, GetHostServices(), L"builtin/file-system-mtp-other", invalidIdFs.put_void());
    state.Require(invalidIdHr == HRESULT_FROM_WIN32(ERROR_NOT_FOUND) && ! invalidIdFs,
                  std::format(L"MTP factory: invalid id expected ERROR_NOT_FOUND/null, got hr=0x{:08X}.", static_cast<unsigned long>(invalidIdHr)));

    void* wrongInterface      = nullptr;
    const HRESULT wrongRiidHr = createFactory(__uuidof(IFileSystemIO), &factoryOptions, GetHostServices(), L"builtin/file-system-mtp", &wrongInterface);
    state.Require(wrongRiidHr == E_NOINTERFACE && wrongInterface == nullptr,
                  std::format(L"MTP factory: wrong riid expected E_NOINTERFACE/null, got hr=0x{:08X}.", static_cast<unsigned long>(wrongRiidHr)));

    const HRESULT nullResultHr = createFactory(__uuidof(IFileSystem), &factoryOptions, GetHostServices(), L"builtin/file-system-mtp", nullptr);
    state.Require(nullResultHr == E_POINTER,
                  std::format(L"MTP factory: null result expected E_POINTER, got hr=0x{:08X}.", static_cast<unsigned long>(nullResultHr)));

    const auto requireSchemaAccepted = [&](const wchar_t* pluginId, std::wstring_view label) noexcept
    {
        const char* schema = nullptr;
        const HRESULT hr   = getSchema(__uuidof(IFileSystem), pluginId, &schema);
        state.Require(SUCCEEDED(hr) && schema != nullptr && schema[0] != '\0',
                      std::format(L"MTP factory: {} schema expected success/non-empty, got hr=0x{:08X}.", label, static_cast<unsigned long>(hr)));
    };

    requireSchemaAccepted(nullptr, L"nullptr id");
    requireSchemaAccepted(L"", L"empty id");
    requireSchemaAccepted(L"builtin/file-system-mtp", L"canonical id");

    const char* invalidSchema     = nullptr;
    const HRESULT invalidSchemaHr = getSchema(__uuidof(IFileSystem), L"builtin/file-system-mtp-other", &invalidSchema);
    state.Require(invalidSchemaHr == HRESULT_FROM_WIN32(ERROR_NOT_FOUND) && invalidSchema == nullptr,
                  std::format(L"MTP factory: invalid id schema expected ERROR_NOT_FOUND/null, got hr=0x{:08X}.", static_cast<unsigned long>(invalidSchemaHr)));

    return state.failure.empty();
});

SelfTest::RunCase(options,
                  suite,
                  L"mtp_live_device_smoke",
                  [&](SelfTest::CaseState& state) noexcept
{
    const std::wstring requestedDevice = GetEnvVarTrimmed(L"REDSALAMANDER_SELFTEST_MTP_DEVICE");
    if (requestedDevice.empty())
    {
        return state.Skip(L"MTP live smoke skipped: set REDSALAMANDER_SELFTEST_MTP_DEVICE to run against an approved device.");
    }

    CreatedFileSystemInstance created;
    const HRESULT createHr = TryCreateFileSystemInstanceWithHost(kBuiltinMtpFileSystemId, GetHostServices(), L"/", created);
    state.Require(SUCCEEDED(createHr) && created.fileSystem,
                  std::format(L"MTP live smoke: create production instance failed. hr=0x{:08X}", static_cast<unsigned long>(createHr)));
    if (FAILED(createHr) || ! created.fileSystem)
    {
        return false;
    }

    wil::com_ptr<IInformations> informations;
    const HRESULT qiInfos = created.fileSystem->QueryInterface(__uuidof(IInformations), informations.put_void());
    state.Require(SUCCEEDED(qiInfos) && informations, std::format(L"MTP live smoke: missing IInformations. hr=0x{:08X}", static_cast<unsigned long>(qiInfos)));
    if (FAILED(qiInfos) || ! informations)
    {
        return false;
    }

    const HRESULT configHr = informations->SetConfiguration(R"json({"readOnly":true,"commandTimeoutMs":15000})json");
    state.Require(SUCCEEDED(configHr), std::format(L"MTP live smoke: SetConfiguration failed. hr=0x{:08X}", static_cast<unsigned long>(configHr)));
    if (FAILED(configHr))
    {
        return false;
    }

    const auto rootEntries = SnapshotDirectoryEntries(created.fileSystem, L"/", state, L"MTP live smoke root");
    if (! state.failure.empty())
    {
        return false;
    }
    if (rootEntries.empty())
    {
        return state.Skip(L"MTP live smoke skipped: no WPD/MTP devices enumerated.");
    }

    for (const DirectoryEntrySnapshot& entry : rootEntries)
    {
        state.Require((entry.attributes & FILE_ATTRIBUTE_DIRECTORY) != 0,
                      std::format(L"MTP live smoke: root entry {} was not reported as a directory.", entry.name));
    }

    const DirectoryEntrySnapshot* selectedDevice = nullptr;
    if (requestedDevice == L"*")
    {
        selectedDevice = &rootEntries.front();
    }
    else
    {
        selectedDevice = FindDirectoryEntrySnapshot(rootEntries, requestedDevice);
    }
    if (selectedDevice == nullptr)
    {
        return state.Skip(std::format(L"MTP live smoke skipped: requested device '{}' was not found in WPD root.", requestedDevice));
    }

    const std::wstring deviceRoot = JoinPluginPathForSelfTest(L"/", selectedDevice->name);
    const auto storageEntries     = SnapshotDirectoryEntries(created.fileSystem, deviceRoot.c_str(), state, L"MTP live smoke device root");
    if (! state.failure.empty())
    {
        return false;
    }

    const DirectoryEntrySnapshot* selectedStorage = nullptr;
    for (const DirectoryEntrySnapshot& entry : storageEntries)
    {
        if ((entry.attributes & FILE_ATTRIBUTE_DIRECTORY) != 0)
        {
            selectedStorage = &entry;
            break;
        }
    }
    if (selectedStorage == nullptr)
    {
        return state.Skip(std::format(L"MTP live smoke skipped: device '{}' exposed no writable storage roots.", selectedDevice->name));
    }

    wil::com_ptr<IFileSystemIO> io;
    wil::com_ptr<IFileSystemDirectoryOperations> dirOps;
    state.Require(CreateFileSystemIo(created.fileSystem, io), L"MTP live smoke: missing IFileSystemIO.");
    state.Require(CreateFileSystemDirectoryOperations(created.fileSystem, dirOps), L"MTP live smoke: missing IFileSystemDirectoryOperations.");
    if (! io || ! dirOps)
    {
        return false;
    }

    const HRESULT writableConfigHr =
        informations->SetConfiguration(R"json({"readOnly":false,"commandTimeoutMs":30000,"byteVerifyOnOverwrite":"deviceReread"})json");
    state.Require(SUCCEEDED(writableConfigHr),
                  std::format(L"MTP live smoke: writable SetConfiguration failed. hr=0x{:08X}", static_cast<unsigned long>(writableConfigHr)));
    if (FAILED(writableConfigHr))
    {
        return false;
    }

    const std::wstring storageRoot = JoinPluginPathForSelfTest(deviceRoot, selectedStorage->name);
    const std::wstring guid        = MakeGuidText();
    state.Require(! guid.empty(), L"MTP live smoke: failed to generate unique scratch name.");
    if (guid.empty())
    {
        return false;
    }

    const std::wstring requestedScratch = GetEnvVarTrimmed(L"REDSALAMANDER_SELFTEST_MTP_SCRATCH");
    std::wstring scratchParent          = storageRoot;
    bool createdScratchParent           = false;
    if (! requestedScratch.empty())
    {
        scratchParent = (requestedScratch.front() == L'/' || requestedScratch.front() == L'\\') ? NormalizePluginPathForSelfTest(requestedScratch)
                                                                                                : JoinPluginPathForSelfTest(storageRoot, requestedScratch);

        unsigned long scratchAttrs  = 0;
        const HRESULT scratchAttrHr = io->GetAttributes(scratchParent.c_str(), &scratchAttrs);
        if (scratchAttrHr == HRESULT_FROM_WIN32(ERROR_FILE_NOT_FOUND) || scratchAttrHr == HRESULT_FROM_WIN32(ERROR_PATH_NOT_FOUND))
        {
            const HRESULT createScratchHr = dirOps->CreateDirectory(scratchParent.c_str());
            state.Require(
                SUCCEEDED(createScratchHr),
                std::format(L"MTP live smoke: CreateDirectory('{}') failed. hr=0x{:08X}", scratchParent, static_cast<unsigned long>(createScratchHr)));
            if (FAILED(createScratchHr))
            {
                return false;
            }
            createdScratchParent = true;
        }
        else
        {
            state.Require(
                SUCCEEDED(scratchAttrHr) && (scratchAttrs & FILE_ATTRIBUTE_DIRECTORY) != 0,
                std::format(L"MTP live smoke: scratch path '{}' is not a directory. hr=0x{:08X}", scratchParent, static_cast<unsigned long>(scratchAttrHr)));
            if (FAILED(scratchAttrHr) || (scratchAttrs & FILE_ATTRIBUTE_DIRECTORY) == 0)
            {
                return false;
            }

            const auto scratchEntries = SnapshotDirectoryEntries(created.fileSystem, scratchParent.c_str(), state, L"MTP live smoke scratch parent");
            if (! state.failure.empty())
            {
                return false;
            }
            if (! scratchEntries.empty())
            {
                return state.Skip(std::format(L"MTP live smoke skipped: requested scratch path '{}' is not empty.", scratchParent));
            }
        }
    }

    const std::wstring scratchRoot =
        JoinPluginPathForSelfTest(scratchParent, requestedScratch.empty() ? std::format(L".RedSalamanderMtpSmoke-{}", guid) : std::format(L"live-{}", guid));
    const HRESULT mkdirHr = dirOps->CreateDirectory(scratchRoot.c_str());
    state.Require(SUCCEEDED(mkdirHr),
                  std::format(L"MTP live smoke: CreateDirectory('{}') failed. hr=0x{:08X}", scratchRoot, static_cast<unsigned long>(mkdirHr)));
    if (FAILED(mkdirHr))
    {
        return false;
    }

    auto cleanup = wil::scope_exit([&]() noexcept
    {
        static_cast<void>(created.fileSystem->DeleteItem(scratchRoot.c_str(), FILESYSTEM_FLAG_RECURSIVE, nullptr, nullptr, nullptr));
        if (createdScratchParent)
        {
            static_cast<void>(created.fileSystem->DeleteItem(scratchParent.c_str(), FILESYSTEM_FLAG_NONE, nullptr, nullptr, nullptr));
        }
    });

    std::string readBack;
    constexpr std::string_view kPayload          = "RedSalamander live MTP smoke payload\n";
    constexpr std::string_view kOverwritePayload = "RedSalamander live MTP smoke replacement\n";
    const std::wstring roundtripPath             = JoinPluginPathForSelfTest(scratchRoot, L"roundtrip.bin");
    state.Require(WritePluginFileText(io.get(), roundtripPath.c_str(), FILESYSTEM_FLAG_NONE, kPayload, state, L"MTP live smoke write"),
                  L"MTP live smoke: failed to write round-trip file.");
    state.Require(ReadPluginFileText(io.get(), roundtripPath.c_str(), readBack, state, L"MTP live smoke read"),
                  L"MTP live smoke: failed to read round-trip file.");
    state.Require(readBack == kPayload, L"MTP live smoke: round-trip file contents did not match.");

    state.Require(WritePluginFileText(io.get(), roundtripPath.c_str(), FILESYSTEM_FLAG_ALLOW_OVERWRITE, kOverwritePayload, state, L"MTP live smoke overwrite"),
                  L"MTP live smoke: failed to overwrite round-trip file.");
    state.Require(ReadPluginFileText(io.get(), roundtripPath.c_str(), readBack, state, L"MTP live smoke overwrite read"),
                  L"MTP live smoke: failed to read overwritten file.");
    state.Require(readBack == kOverwritePayload, L"MTP live smoke: overwritten file contents did not match.");

    const std::wstring renamedPath = JoinPluginPathForSelfTest(scratchRoot, L"renamed.bin");
    const HRESULT renameHr = created.fileSystem->RenameItem(roundtripPath.c_str(), renamedPath.c_str(), FILESYSTEM_FLAG_NONE, nullptr, nullptr, nullptr);
    state.Require(SUCCEEDED(renameHr), std::format(L"MTP live smoke: RenameItem failed. hr=0x{:08X}", static_cast<unsigned long>(renameHr)));
    state.Require(ReadPluginFileText(io.get(), renamedPath.c_str(), readBack, state, L"MTP live smoke renamed read"),
                  L"MTP live smoke: failed to read renamed file.");
    state.Require(readBack == kOverwritePayload, L"MTP live smoke: renamed file contents did not match.");

    const std::wstring copiedPath = JoinPluginPathForSelfTest(scratchRoot, L"copied.bin");
    const HRESULT copyHr          = created.fileSystem->CopyItem(renamedPath.c_str(), copiedPath.c_str(), FILESYSTEM_FLAG_NONE, nullptr, nullptr, nullptr);
    state.Require(SUCCEEDED(copyHr), std::format(L"MTP live smoke: CopyItem failed. hr=0x{:08X}", static_cast<unsigned long>(copyHr)));
    state.Require(ReadPluginFileText(io.get(), copiedPath.c_str(), readBack, state, L"MTP live smoke copied read"),
                  L"MTP live smoke: failed to read copied file.");
    state.Require(readBack == kOverwritePayload, L"MTP live smoke: copied file contents did not match.");

    const std::wstring movedPath = JoinPluginPathForSelfTest(scratchRoot, L"moved.bin");
    const HRESULT moveHr         = created.fileSystem->MoveItem(copiedPath.c_str(), movedPath.c_str(), FILESYSTEM_FLAG_NONE, nullptr, nullptr, nullptr);
    state.Require(SUCCEEDED(moveHr), std::format(L"MTP live smoke: MoveItem failed. hr=0x{:08X}", static_cast<unsigned long>(moveHr)));
    unsigned long attrs = 0;
    state.Require(io->GetAttributes(copiedPath.c_str(), &attrs) == HRESULT_FROM_WIN32(ERROR_FILE_NOT_FOUND), L"MTP live smoke: MoveItem left source behind.");
    state.Require(SUCCEEDED(io->GetAttributes(movedPath.c_str(), &attrs)) && (attrs & FILE_ATTRIBUTE_DIRECTORY) == 0,
                  L"MTP live smoke: moved file is not readable as a file.");

    const HRESULT deleteMovedHr = created.fileSystem->DeleteItem(movedPath.c_str(), FILESYSTEM_FLAG_NONE, nullptr, nullptr, nullptr);
    state.Require(SUCCEEDED(deleteMovedHr), std::format(L"MTP live smoke: DeleteItem(file) failed. hr=0x{:08X}", static_cast<unsigned long>(deleteMovedHr)));

    const HRESULT deleteRootHr = created.fileSystem->DeleteItem(scratchRoot.c_str(), FILESYSTEM_FLAG_RECURSIVE, nullptr, nullptr, nullptr);
    state.Require(SUCCEEDED(deleteRootHr),
                  std::format(L"MTP live smoke: DeleteItem(scratch root) failed. hr=0x{:08X}", static_cast<unsigned long>(deleteRootHr)));
    if (SUCCEEDED(deleteRootHr))
    {
        cleanup.release();
        if (createdScratchParent)
        {
            static_cast<void>(created.fileSystem->DeleteItem(scratchParent.c_str(), FILESYSTEM_FLAG_NONE, nullptr, nullptr, nullptr));
        }
    }

    return state.failure.empty();
});

SelfTest::RunCase(options,
                  suite,
                  L"mtp_queryinterface_matrix",
                  [&](SelfTest::CaseState& state) noexcept
{
    CreatedFileSystemInstance created;
    const HRESULT createHr = TryCreateFakeMtpFileSystemInstance("{}", R"json({"readOnly":false})json", L"/", created);
    if (createHr == HRESULT_FROM_WIN32(ERROR_PROC_NOT_FOUND))
    {
        return state.Skip(L"MTP fake backend export is available only in debug plugin builds.");
    }
    state.Require(SUCCEEDED(createHr) && created.fileSystem,
                  std::format(L"MTP QI: create selftest instance failed. hr=0x{:08X}", static_cast<unsigned long>(createHr)));
    if (FAILED(createHr) || ! created.fileSystem)
    {
        return false;
    }

    wil::com_ptr<IUnknown> canonicalUnknown;
    const HRESULT unknownHr = created.fileSystem->QueryInterface(__uuidof(IUnknown), canonicalUnknown.put_void());
    state.Require(SUCCEEDED(unknownHr) && canonicalUnknown, std::format(L"MTP QI: IUnknown query failed. hr=0x{:08X}", static_cast<unsigned long>(unknownHr)));
    if (FAILED(unknownHr) || ! canonicalUnknown)
    {
        return false;
    }

    RequireMtpSupportedQiIdentity<IFileSystem>(created.fileSystem.get(), canonicalUnknown.get(), state, L"IFileSystem");
    RequireMtpSupportedQiIdentity<IFileSystemIO>(created.fileSystem.get(), canonicalUnknown.get(), state, L"IFileSystemIO");
    RequireMtpSupportedQiIdentity<IFileSystemDirectoryOperations>(created.fileSystem.get(), canonicalUnknown.get(), state, L"IFileSystemDirectoryOperations");
    RequireMtpSupportedQiIdentity<IFileSystemInitialize>(created.fileSystem.get(), canonicalUnknown.get(), state, L"IFileSystemInitialize");
    RequireMtpSupportedQiIdentity<IInformations>(created.fileSystem.get(), canonicalUnknown.get(), state, L"IInformations");
    RequireMtpSupportedQiIdentity<INavigationMenu>(created.fileSystem.get(), canonicalUnknown.get(), state, L"INavigationMenu");
    RequireMtpSupportedQiIdentity<IDriveInfo>(created.fileSystem.get(), canonicalUnknown.get(), state, L"IDriveInfo");

    RequireMtpUnsupportedQi(created.fileSystem.get(), __uuidof(IFileSystemDirectoryWatch), state, L"IFileSystemDirectoryWatch");
    RequireMtpUnsupportedQi(created.fileSystem.get(), __uuidof(IFileSystemSearch), state, L"IFileSystemSearch");
    RequireMtpUnsupportedQi(created.fileSystem.get(), __uuidof(IFileSystemItemStreams), state, L"IFileSystemItemStreams");
    RequireMtpUnsupportedQi(created.fileSystem.get(), __uuidof(IFilesInformation), state, L"IFilesInformation");
    RequireMtpUnsupportedQi(created.fileSystem.get(), __uuidof(IFileReader), state, L"IFileReader");
    RequireMtpUnsupportedQi(created.fileSystem.get(), __uuidof(IFileWriter), state, L"IFileWriter");

    const HRESULT nullResultHr = created.fileSystem->QueryInterface(__uuidof(IFileSystemIO), nullptr);
    state.Require(nullResultHr == E_POINTER,
                  std::format(L"MTP QI: null result expected E_POINTER, got hr=0x{:08X}.", static_cast<unsigned long>(nullResultHr)));
    return state.failure.empty();
});

SelfTest::RunCase(options,
                  suite,
                  L"mtp_json_return_buffers_are_bounded",
                  [&](SelfTest::CaseState& state) noexcept
{
    CreatedFileSystemInstance created;
    const HRESULT createHr = TryCreateFakeMtpFileSystemInstance("{}", R"json({"readOnly":false,"byteVerifyOnOverwrite":"deviceReread"})json", L"/", created);
    if (createHr == HRESULT_FROM_WIN32(ERROR_PROC_NOT_FOUND))
    {
        return state.Skip(L"MTP fake backend export is available only in debug plugin builds.");
    }
    state.Require(SUCCEEDED(createHr) && created.fileSystem,
                  std::format(L"MTP JSON buffers: create selftest instance failed. hr=0x{:08X}", static_cast<unsigned long>(createHr)));
    if (FAILED(createHr) || ! created.fileSystem)
    {
        return false;
    }

    wil::com_ptr<IInformations> informations;
    wil::com_ptr<IFileSystemIO> io;
    state.Require(CreateInformations(created.fileSystem, informations), L"MTP JSON buffers: missing IInformations.");
    state.Require(CreateFileSystemIo(created.fileSystem, io), L"MTP JSON buffers: missing IFileSystemIO.");
    if (! informations || ! io)
    {
        return false;
    }

    const auto uniquePointerCount = [](const std::vector<const char*>& values) noexcept
    {
        std::vector<const char*> uniqueValues;
        uniqueValues.reserve(values.size());
        for (const char* value : values)
        {
            const auto it = std::find(uniqueValues.begin(), uniqueValues.end(), value);
            if (it == uniqueValues.end())
            {
                uniqueValues.push_back(value);
            }
        }
        return uniqueValues.size();
    };

    const auto requireTwoSlotMethod = [&](std::wstring_view label, auto&& getJson, std::string_view expectedToken) noexcept
    {
        std::vector<const char*> pointers;
        pointers.reserve(8u);
        for (size_t i = 0; i < 8u; ++i)
        {
            const char* json = nullptr;
            const HRESULT hr = getJson(&json);
            state.Require(SUCCEEDED(hr) && json != nullptr && json[0] != '\0',
                          std::format(L"MTP JSON buffers: {} call {} failed. hr=0x{:08X}", label, i, static_cast<unsigned long>(hr)));
            if (FAILED(hr) || json == nullptr)
            {
                return;
            }

            if (i < 2u)
            {
                state.Require(std::string_view(json).find(expectedToken) != std::string_view::npos,
                              std::format(L"MTP JSON buffers: {} omitted expected token.", label));
            }
            pointers.push_back(json);
        }

        const size_t uniqueCount = uniquePointerCount(pointers);
        state.Require(uniqueCount <= 2u, std::format(L"MTP JSON buffers: {} returned {} unique pointers; expected at most 2.", label, uniqueCount));
        state.Require(pointers.size() < 3u || pointers[0] == pointers[2],
                      std::format(L"MTP JSON buffers: {} did not rotate back to slot 0 on the third call.", label));
        state.Require(pointers.size() < 4u || pointers[1] == pointers[3],
                      std::format(L"MTP JSON buffers: {} did not rotate back to slot 1 on the fourth call.", label));
    };

    requireTwoSlotMethod(L"GetConfiguration", [&](const char** json) noexcept {
        return informations->GetConfiguration(json);
    }, R"json("byteVerifyOnOverwrite":"deviceReread")json");
    requireTwoSlotMethod(L"GetCapabilities", [&](const char** json) noexcept {
        return created.fileSystem->GetCapabilities(json);
    }, R"json("byteVerifyOnOverwrite": "deviceReread")json");
    requireTwoSlotMethod(L"GetItemProperties", [&](const char** json) noexcept {
        return io->GetItemProperties(L"/Fake Phone [devpuid:fake-device]/Internal Storage/DCIM/Camera/photo001.txt", json);
    }, R"json("persistentId":"file-photo001")json");

    return state.failure.empty();
});

SelfTest::RunCase(options,
                  suite,
                  L"mtp_capabilities_are_instance_honest",
                  [&](SelfTest::CaseState& state) noexcept
{
    const auto requireCapabilityToken = [&](IFileSystem* fs, std::string_view token, std::wstring_view label) noexcept
    {
        if (fs == nullptr)
        {
            state.Require(false, std::format(L"MTP capabilities: {} filesystem is null.", label));
            return;
        }

        const char* json = nullptr;
        const HRESULT hr = fs->GetCapabilities(&json);
        state.Require(SUCCEEDED(hr) && json != nullptr && json[0] != '\0',
                      std::format(L"MTP capabilities: {} GetCapabilities failed. hr=0x{:08X}", label, static_cast<unsigned long>(hr)));
        if (FAILED(hr) || json == nullptr)
        {
            return;
        }

        const std::string_view capabilities(json);
        state.Require(capabilities.find(token) != std::string_view::npos,
                      std::format(L"MTP capabilities: {} omitted expected token '{}'.", label, std::wstring(token.begin(), token.end())));
        state.Require(capabilities.find(R"json("version": 1)json") != std::string_view::npos, std::format(L"MTP capabilities: {} omitted version.", label));
        state.Require(capabilities.find(R"json("operations")json") != std::string_view::npos, std::format(L"MTP capabilities: {} omitted operations.", label));
        state.Require(capabilities.find(R"json("concurrency")json") != std::string_view::npos,
                      std::format(L"MTP capabilities: {} omitted concurrency.", label));
        state.Require(capabilities.find(R"json("crossFileSystem")json") != std::string_view::npos,
                      std::format(L"MTP capabilities: {} omitted crossFileSystem.", label));
        state.Require(capabilities.find(R"json("pathIdentity")json") != std::string_view::npos,
                      std::format(L"MTP capabilities: {} omitted pathIdentity.", label));
    };

    CreatedFileSystemInstance production;
    const HRESULT productionHr = TryCreateFileSystemInstance(kBuiltinMtpFileSystemId, L"", production);
    state.Require(SUCCEEDED(productionHr) && production.fileSystem,
                  std::format(L"MTP capabilities: production create failed. hr=0x{:08X}", static_cast<unsigned long>(productionHr)));
    if (SUCCEEDED(productionHr) && production.fileSystem)
    {
        requireCapabilityToken(production.fileSystem.get(), R"json("backend": "wpd")json", L"production");
        requireCapabilityToken(production.fileSystem.get(), R"json("write": false)json", L"production");
    }

    CreatedFileSystemInstance writableFake;
    const HRESULT writableHr = TryCreateFakeMtpFileSystemInstance("{}", R"json({"readOnly":false,"byteVerifyOnOverwrite":"sizeOnly"})json", L"/", writableFake);
    if (writableHr == HRESULT_FROM_WIN32(ERROR_PROC_NOT_FOUND))
    {
        return state.Skip(L"MTP fake backend export is available only in debug plugin builds.");
    }
    state.Require(SUCCEEDED(writableHr) && writableFake.fileSystem,
                  std::format(L"MTP capabilities: writable fake create failed. hr=0x{:08X}", static_cast<unsigned long>(writableHr)));
    if (SUCCEEDED(writableHr) && writableFake.fileSystem)
    {
        requireCapabilityToken(writableFake.fileSystem.get(), R"json("backend": "fake")json", L"writable fake");
        requireCapabilityToken(writableFake.fileSystem.get(), R"json("write": true)json", L"writable fake");
        requireCapabilityToken(writableFake.fileSystem.get(), R"json("copy": true)json", L"writable fake");
        requireCapabilityToken(writableFake.fileSystem.get(), R"json("move": true)json", L"writable fake");
        requireCapabilityToken(writableFake.fileSystem.get(), R"json("delete": true)json", L"writable fake");
        requireCapabilityToken(writableFake.fileSystem.get(), R"json("rename": true)json", L"writable fake");
        requireCapabilityToken(writableFake.fileSystem.get(), R"json("byteVerifyOnOverwrite": "sizeOnly")json", L"writable fake");
    }

    CreatedFileSystemInstance readOnlyFake;
    const HRESULT readOnlyHr =
        TryCreateFakeMtpFileSystemInstance("{}", R"json({"readOnly":true,"byteVerifyOnOverwrite":"deviceReread"})json", L"/", readOnlyFake);
    state.Require(SUCCEEDED(readOnlyHr) && readOnlyFake.fileSystem,
                  std::format(L"MTP capabilities: read-only fake create failed. hr=0x{:08X}", static_cast<unsigned long>(readOnlyHr)));
    if (SUCCEEDED(readOnlyHr) && readOnlyFake.fileSystem)
    {
        requireCapabilityToken(readOnlyFake.fileSystem.get(), R"json("backend": "fake")json", L"read-only fake");
        requireCapabilityToken(readOnlyFake.fileSystem.get(), R"json("write": false)json", L"read-only fake");
        requireCapabilityToken(readOnlyFake.fileSystem.get(), R"json("copy": false)json", L"read-only fake");
        requireCapabilityToken(readOnlyFake.fileSystem.get(), R"json("move": false)json", L"read-only fake");
        requireCapabilityToken(readOnlyFake.fileSystem.get(), R"json("delete": false)json", L"read-only fake");
        requireCapabilityToken(readOnlyFake.fileSystem.get(), R"json("rename": false)json", L"read-only fake");
        requireCapabilityToken(readOnlyFake.fileSystem.get(), R"json("byteVerifyOnOverwrite": "deviceReread")json", L"read-only fake");
    }

    return state.failure.empty();
});

SelfTest::RunCase(options,
                  suite,
                  L"mtp_path_scheme_and_device_key_normalization",
                  [&](SelfTest::CaseState& state) noexcept
{
    constexpr std::wstring_view kDcimRoot            = L"/Fake Phone [devpuid:fake-device]/Internal Storage/DCIM";
    const std::array<std::wstring_view, 4> rootForms = {
        kDcimRoot,
        L"mtp:/Fake Phone [devpuid:fake-device]/Internal Storage/DCIM",
        L"mtp://Fake Phone [devpuid:fake-device]/Internal Storage/DCIM",
        L"\\Fake Phone [devpuid:fake-device]\\Internal Storage\\DCIM\\",
    };

    for (const std::wstring_view rootForm : rootForms)
    {
        CreatedFileSystemInstance created;
        const HRESULT createHr = TryCreateFakeMtpFileSystemInstance("{}", R"json({"readOnly":false})json", rootForm, created);
        if (createHr == HRESULT_FROM_WIN32(ERROR_PROC_NOT_FOUND))
        {
            return state.Skip(L"MTP fake backend export is available only in debug plugin builds.");
        }
        state.Require(SUCCEEDED(createHr) && created.fileSystem,
                      std::format(L"MTP path normalization: create failed for root '{}'. hr=0x{:08X}", rootForm, static_cast<unsigned long>(createHr)));
        if (FAILED(createHr) || ! created.fileSystem)
        {
            return false;
        }

        const auto rootEntries = SnapshotDirectoryEntries(created.fileSystem, L"/", state, L"MTP path normalization root");
        state.Require(FindDirectoryEntrySnapshot(rootEntries, L"Camera") != nullptr,
                      std::format(L"MTP path normalization: '/' did not map to DCIM root for '{}'.", rootForm));

        const auto cameraEntries = SnapshotDirectoryEntries(created.fileSystem, L"mtp:/Camera", state, L"MTP path normalization camera");
        state.Require(FindDirectoryEntrySnapshot(cameraEntries, L"photo001.txt") != nullptr,
                      std::format(L"MTP path normalization: 'mtp:/Camera' did not map under root '{}'.", rootForm));

        wil::com_ptr<IFileSystemIO> io;
        state.Require(CreateFileSystemIo(created.fileSystem, io), L"MTP path normalization: missing IFileSystemIO.");
        if (! io)
        {
            return false;
        }

        std::string readBack;
        state.Require(ReadPluginFileText(io.get(), L"/@conn:Any Saved Profile/Camera/photo001.txt", readBack, state, L"MTP path normalization conn read"),
                      L"MTP path normalization: failed to read via /@conn: suffix path.");
        state.Require(readBack == "RedSalamander deterministic MTP fixture\r\n", L"MTP path normalization: /@conn: suffix path read returned wrong contents.");

        readBack.clear();
        state.Require(ReadPluginFileText(io.get(), L"mtp://Camera/photo001.txt", readBack, state, L"MTP path normalization scheme read"),
                      L"MTP path normalization: failed to read via mtp:// relative path.");
        state.Require(readBack == "RedSalamander deterministic MTP fixture\r\n", L"MTP path normalization: mtp:// relative path read returned wrong contents.");
    }

    return state.failure.empty();
});

SelfTest::RunCase(options,
                  suite,
                  L"mtp_duplicate_names_require_stable_suffix",
                  [&](SelfTest::CaseState& state) noexcept
{
    CreatedFileSystemInstance created;
    const HRESULT createHr = TryCreateFakeMtpFileSystemInstance("{}", R"json({"readOnly":false})json", L"/", created);
    if (createHr == HRESULT_FROM_WIN32(ERROR_PROC_NOT_FOUND))
    {
        return state.Skip(L"MTP fake backend export is available only in debug plugin builds.");
    }
    state.Require(SUCCEEDED(createHr) && created.fileSystem,
                  std::format(L"MTP duplicate names: create selftest instance failed. hr=0x{:08X}", static_cast<unsigned long>(createHr)));
    if (FAILED(createHr) || ! created.fileSystem)
    {
        return false;
    }

    wil::com_ptr<IFileSystemIO> io;
    state.Require(CreateFileSystemIo(created.fileSystem, io), L"MTP duplicate names: missing IFileSystemIO.");
    if (! io)
    {
        return false;
    }

    constexpr std::wstring_view kCameraPath = L"/Fake Phone [devpuid:fake-device]/Internal Storage/DCIM/Camera";
    const auto cameraEntries                = SnapshotDirectoryEntries(created.fileSystem, kCameraPath.data(), state, L"MTP duplicate names camera");

    constexpr std::wstring_view kDuplicatePrefix = L"duplicate-sibling.txt [puid:";
    std::vector<std::wstring> duplicateNames;
    for (const DirectoryEntrySnapshot& entry : cameraEntries)
    {
        if (entry.name.rfind(kDuplicatePrefix, 0) == 0)
        {
            duplicateNames.push_back(entry.name);
        }
    }
    std::sort(duplicateNames.begin(), duplicateNames.end());

    state.Require(duplicateNames.size() == 2u, std::format(L"MTP duplicate names: expected two suffixed duplicate entries, got {}.", duplicateNames.size()));
    state.Require(FindDirectoryEntrySnapshot(cameraEntries, L"duplicate-sibling.txt") == nullptr,
                  L"MTP duplicate names: unsuffixed ambiguous duplicate entry was exposed.");

    for (const std::wstring& duplicateName : duplicateNames)
    {
        const bool suffixShapeOk = duplicateName.size() == kDuplicatePrefix.size() + 16u + 1u && duplicateName.back() == L']' &&
                                   std::all_of(duplicateName.begin() + static_cast<std::ptrdiff_t>(kDuplicatePrefix.size()),
                                               duplicateName.end() - 1,
                                               [](wchar_t ch) noexcept { return (ch >= L'0' && ch <= L'9') || (ch >= L'A' && ch <= L'F'); });
        state.Require(suffixShapeOk, std::format(L"MTP duplicate names: '{}' does not use the expected [puid:HEX] suffix.", duplicateName));
    }

    std::vector<std::string> duplicateContents;
    for (const std::wstring& duplicateName : duplicateNames)
    {
        const std::wstring path = std::wstring(kCameraPath) + L"/" + duplicateName;
        std::string readBack;
        state.Require(ReadPluginFileText(io.get(), path.c_str(), readBack, state, L"MTP duplicate names duplicate read"),
                      std::format(L"MTP duplicate names: failed to read suffixed duplicate '{}'.", duplicateName));
        duplicateContents.push_back(std::move(readBack));
    }
    std::sort(duplicateContents.begin(), duplicateContents.end());
    state.Require(duplicateContents.size() == 2u && duplicateContents[0] == "duplicate sibling one\r\n" && duplicateContents[1] == "duplicate sibling two\r\n",
                  L"MTP duplicate names: suffixed duplicate reads did not return the expected distinct contents.");

    constexpr std::array<std::wstring_view, 5> kLiteralNames = {{
        L"name [puid:literal].txt",
        L"percent %.txt",
        L"bracket ].txt",
        L"trailing-space .txt",
        L"caf\u00E9.txt",
    }};
    for (const std::wstring_view literalName : kLiteralNames)
    {
        state.Require(FindDirectoryEntrySnapshot(cameraEntries, literalName) != nullptr,
                      std::format(L"MTP duplicate names: literal/special name '{}' was not preserved.", literalName));
    }

    const std::wstring ambiguousPath = std::wstring(kCameraPath) + L"/duplicate-sibling.txt";
    const HRESULT deleteHr           = created.fileSystem->DeleteItem(ambiguousPath.c_str(), FILESYSTEM_FLAG_NONE, nullptr, nullptr, nullptr);
    state.Require(deleteHr == HRESULT_FROM_WIN32(ERROR_FILE_NOT_FOUND),
                  std::format(L"MTP duplicate names: unsuffixed ambiguous delete should fail closed, got hr=0x{:08X}.", static_cast<unsigned long>(deleteHr)));

    const auto afterDeleteAttempt = SnapshotDirectoryEntries(created.fileSystem, kCameraPath.data(), state, L"MTP duplicate names after ambiguous delete");
    std::vector<std::wstring> afterDuplicateNames;
    for (const DirectoryEntrySnapshot& entry : afterDeleteAttempt)
    {
        if (entry.name.rfind(kDuplicatePrefix, 0) == 0)
        {
            afterDuplicateNames.push_back(entry.name);
        }
    }
    std::sort(afterDuplicateNames.begin(), afterDuplicateNames.end());
    state.Require(afterDuplicateNames == duplicateNames, L"MTP duplicate names: failed ambiguous mutation changed duplicate entries.");

    return state.failure.empty();
});

SelfTest::RunCase(options,
                  suite,
                  L"mtp_disconnect_mid_enumeration_surfaces_error",
                  [&](SelfTest::CaseState& state) noexcept
{
    constexpr std::wstring_view kCameraPath = L"/Fake Phone [devpuid:fake-device]/Internal Storage/DCIM/Camera";

    CreatedFileSystemInstance created;
    const HRESULT createHr =
        TryCreateFakeMtpFileSystemInstance(R"json({"disconnectEnumerateOncePath":"/Fake Phone [devpuid:fake-device]/Internal Storage/DCIM/Camera"})json",
                                           R"json({"readOnly":true})json",
                                           L"/",
                                           created);
    if (createHr == HRESULT_FROM_WIN32(ERROR_PROC_NOT_FOUND))
    {
        return state.Skip(L"MTP fake backend export is available only in debug plugin builds.");
    }
    state.Require(SUCCEEDED(createHr) && created.fileSystem,
                  std::format(L"MTP disconnect enumerate: create selftest instance failed. hr=0x{:08X}", static_cast<unsigned long>(createHr)));
    if (FAILED(createHr) || ! created.fileSystem)
    {
        return false;
    }

    wil::com_ptr<IFilesInformation> failedInfo;
    const auto start                      = std::chrono::steady_clock::now();
    const HRESULT disconnectedHr          = created.fileSystem->ReadDirectoryInfo(kCameraPath.data(), failedInfo.put());
    const auto elapsedMs                  = std::chrono::duration_cast<std::chrono::milliseconds>(std::chrono::steady_clock::now() - start).count();
    constexpr HRESULT kExpectedDisconnect = HRESULT_FROM_WIN32(ERROR_DEVICE_NOT_CONNECTED);
    state.Require(disconnectedHr == kExpectedDisconnect,
                  std::format(L"MTP disconnect enumerate: expected ERROR_DEVICE_NOT_CONNECTED, got 0x{:08X}.", static_cast<unsigned long>(disconnectedHr)));
    state.Require(! failedInfo, L"MTP disconnect enumerate: failed enumeration returned a directory information object.");
    state.Require(elapsedMs < 500, std::format(L"MTP disconnect enumerate: failure took {} ms; expected prompt device-gone return.", elapsedMs));

    const auto recoveredEntries = SnapshotDirectoryEntries(created.fileSystem, kCameraPath.data(), state, L"MTP disconnect enumerate recovery");
    state.Require(FindDirectoryEntrySnapshot(recoveredEntries, L"photo001.txt") != nullptr,
                  L"MTP disconnect enumerate: one-shot disconnect left the directory unrecoverable.");

    return state.failure.empty();
});

SelfTest::RunCase(options,
                  suite,
                  L"mtp_unload_quiet_point_no_callback_after_clear",
                  [&](SelfTest::CaseState& state) noexcept
{
    class BlockingNavigationCallback final : public INavigationMenuCallback
    {
    public:
        BlockingNavigationCallback() = default;

        BlockingNavigationCallback(const BlockingNavigationCallback&)            = delete;
        BlockingNavigationCallback(BlockingNavigationCallback&&)                 = delete;
        BlockingNavigationCallback& operator=(const BlockingNavigationCallback&) = delete;
        BlockingNavigationCallback& operator=(BlockingNavigationCallback&&)      = delete;

        HRESULT STDMETHODCALLTYPE NavigationMenuRequestNavigate(const wchar_t* path, void* cookie) noexcept override
        {
            std::unique_lock lock(_mutex);
            ++_callCount;
            _lastPath   = path ? path : L"";
            _lastCookie = cookie;
            _entered    = true;
            _cv.notify_all();
            _cv.wait(lock, [&]() noexcept { return _release; });
            return S_OK;
        }

        [[nodiscard]] bool WaitForCall(std::chrono::milliseconds timeout) noexcept
        {
            std::unique_lock lock(_mutex);
            return _cv.wait_for(lock, timeout, [&]() noexcept { return _entered; });
        }

        void Release() noexcept
        {
            {
                std::lock_guard lock(_mutex);
                _release = true;
            }
            _cv.notify_all();
        }

        [[nodiscard]] uint32_t CallCount() const
        {
            std::lock_guard lock(_mutex);
            return _callCount;
        }

        [[nodiscard]] std::wstring LastPath() const
        {
            std::lock_guard lock(_mutex);
            return _lastPath;
        }

        [[nodiscard]] void* LastCookie() const
        {
            std::lock_guard lock(_mutex);
            return _lastCookie;
        }

    private:
        mutable std::mutex _mutex;
        std::condition_variable _cv;
        bool _entered       = false;
        bool _release       = false;
        uint32_t _callCount = 0;
        std::wstring _lastPath;
        void* _lastCookie = nullptr;
    };

    CreatedFileSystemInstance created;
    const HRESULT createHr = TryCreateFakeMtpFileSystemInstance("{}", R"json({"readOnly":true})json", L"/", created);
    if (createHr == HRESULT_FROM_WIN32(ERROR_PROC_NOT_FOUND))
    {
        return state.Skip(L"MTP fake backend export is available only in debug plugin builds.");
    }
    state.Require(SUCCEEDED(createHr) && created.fileSystem,
                  std::format(L"MTP callback quiet point: create selftest instance failed. hr=0x{:08X}", static_cast<unsigned long>(createHr)));
    if (FAILED(createHr) || ! created.fileSystem)
    {
        return false;
    }

    wil::com_ptr<INavigationMenu> navigationMenu;
    const HRESULT qiHr = created.fileSystem->QueryInterface(__uuidof(INavigationMenu), navigationMenu.put_void());
    state.Require(SUCCEEDED(qiHr) && navigationMenu,
                  std::format(L"MTP callback quiet point: missing INavigationMenu. hr=0x{:08X}", static_cast<unsigned long>(qiHr)));
    if (FAILED(qiHr) || ! navigationMenu)
    {
        return false;
    }

    const NavigationMenuItem* items = nullptr;
    unsigned int count              = 0;
    const HRESULT menuHr            = navigationMenu->GetMenuItems(&items, &count);
    state.Require(SUCCEEDED(menuHr) && items != nullptr && count > 0u,
                  std::format(L"MTP callback quiet point: GetMenuItems failed or returned no items. hr=0x{:08X}", static_cast<unsigned long>(menuHr)));
    if (FAILED(menuHr) || items == nullptr || count == 0u)
    {
        return false;
    }

    unsigned int commandId = 0;
    std::wstring expectedPath;
    for (unsigned int index = 0; index < count; ++index)
    {
        if (items[index].commandId != 0u && items[index].path != nullptr && items[index].path[0] != L'\0')
        {
            commandId    = items[index].commandId;
            expectedPath = items[index].path;
            break;
        }
    }
    state.Require(commandId != 0u && ! expectedPath.empty(), L"MTP callback quiet point: no navigable menu command was exposed.");
    if (commandId == 0u || expectedPath.empty())
    {
        return false;
    }

    BlockingNavigationCallback callback;
    void* const callbackCookie = &callback;
    HRESULT hr                 = navigationMenu->SetCallback(&callback, callbackCookie);
    state.Require(SUCCEEDED(hr), std::format(L"MTP callback quiet point: SetCallback registration failed. hr=0x{:08X}", static_cast<unsigned long>(hr)));
    if (FAILED(hr))
    {
        return false;
    }

    HRESULT executeHr = E_PENDING;
    std::jthread executeThread([&]() noexcept { executeHr = navigationMenu->ExecuteMenuCommand(commandId); });
    auto joinExecuteThread = wil::scope_exit([&]() noexcept
    {
        callback.Release();
        if (executeThread.joinable())
        {
            executeThread.join();
        }
    });

    state.Require(callback.WaitForCall(std::chrono::seconds(5)), L"MTP callback quiet point: ExecuteMenuCommand did not enter callback.");
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
        callback.Release();
        if (clearThread.joinable())
        {
            clearThread.join();
        }
    });

    const bool clearReturnedWhileCallbackBlocked = [&]() noexcept
    {
        std::unique_lock lock(clearMutex);
        return clearCv.wait_for(lock, std::chrono::milliseconds(200), [&]() noexcept { return clearCompleted; });
    }();
    state.Require(! clearReturnedWhileCallbackBlocked, L"MTP callback quiet point: SetCallback(nullptr, nullptr) returned before the active callback drained.");

    callback.Release();
    const bool clearReturnedAfterRelease = [&]() noexcept
    {
        std::unique_lock lock(clearMutex);
        return clearCv.wait_for(lock, std::chrono::seconds(5), [&]() noexcept { return clearCompleted; });
    }();
    state.Require(clearReturnedAfterRelease, L"MTP callback quiet point: SetCallback(nullptr, nullptr) did not return after callback drained.");

    if (clearThread.joinable())
    {
        clearThread.join();
    }
    if (executeThread.joinable())
    {
        executeThread.join();
    }

    state.Require(SUCCEEDED(clearHr), std::format(L"MTP callback quiet point: clear callback failed. hr=0x{:08X}", static_cast<unsigned long>(clearHr)));
    state.Require(SUCCEEDED(executeHr),
                  std::format(L"MTP callback quiet point: ExecuteMenuCommand failed. hr=0x{:08X}", static_cast<unsigned long>(executeHr)));
    state.Require(callback.CallCount() == 1u,
                  std::format(L"MTP callback quiet point: expected exactly one callback before clear, got {}.", callback.CallCount()));
    state.Require(callback.LastPath() == expectedPath, L"MTP callback quiet point: callback path mismatch before clear.");
    state.Require(callback.LastCookie() == callbackCookie, L"MTP callback quiet point: callback cookie mismatch before clear.");

    const HRESULT staleHr = navigationMenu->ExecuteMenuCommand(commandId);
    state.Require(
        staleHr == HRESULT_FROM_WIN32(ERROR_NOT_FOUND),
        std::format(L"MTP callback quiet point: ExecuteMenuCommand after clear expected ERROR_NOT_FOUND, got 0x{:08X}.", static_cast<unsigned long>(staleHr)));
    state.Require(callback.CallCount() == 1u, L"MTP callback quiet point: callback fired after SetCallback(nullptr, nullptr) returned.");

    return state.failure.empty();
});

SelfTest::RunCase(options,
                  suite,
                  L"mtp_drive_info_and_disconnect_menu",
                  [&](SelfTest::CaseState& state) noexcept
{
    CreatedFileSystemInstance created;
    const HRESULT createHr = TryCreateFakeMtpFileSystemInstance("{}", R"json({"readOnly":true})json", L"/", created);
    if (createHr == HRESULT_FROM_WIN32(ERROR_PROC_NOT_FOUND))
    {
        return state.Skip(L"MTP fake backend export is available only in debug plugin builds.");
    }
    state.Require(SUCCEEDED(createHr) && created.fileSystem,
                  std::format(L"MTP drive info: create selftest instance failed. hr=0x{:08X}", static_cast<unsigned long>(createHr)));
    if (FAILED(createHr) || ! created.fileSystem)
    {
        return false;
    }

    wil::com_ptr<IDriveInfo> driveInfoService;
    const HRESULT qiHr = created.fileSystem->QueryInterface(__uuidof(IDriveInfo), driveInfoService.put_void());
    state.Require(SUCCEEDED(qiHr) && driveInfoService, std::format(L"MTP drive info: missing IDriveInfo. hr=0x{:08X}", static_cast<unsigned long>(qiHr)));
    if (FAILED(qiHr) || ! driveInfoService)
    {
        return false;
    }

    constexpr HRESULT kDeviceDisconnected = HRESULT_FROM_WIN32(ERROR_DEVICE_NOT_CONNECTED);
    DriveInfo drive{};
    HRESULT hr = driveInfoService->GetDriveInfo(L"/Fake Phone [devpuid:fake-device]/Internal Storage/DCIM", &drive);
    state.Require(SUCCEEDED(hr), std::format(L"MTP drive info: GetDriveInfo failed. hr=0x{:08X}", static_cast<unsigned long>(hr)));
    if (FAILED(hr))
    {
        return false;
    }

    constexpr DriveInfoFlags kExpectedFlags =
        static_cast<DriveInfoFlags>(DRIVE_INFO_FLAG_HAS_DISPLAY_NAME | DRIVE_INFO_FLAG_HAS_VOLUME_LABEL | DRIVE_INFO_FLAG_HAS_FILE_SYSTEM);
    state.Require((drive.flags & kExpectedFlags) == kExpectedFlags, L"MTP drive info: missing expected display/volume/file-system flags.");
    state.Require(drive.displayName && std::wstring_view(drive.displayName) == L"Fake Phone [devpuid:fake-device]",
                  L"MTP drive info: display name should use the device root segment.");
    state.Require(drive.volumeLabel && std::wstring_view(drive.volumeLabel) == L"Fake Phone [devpuid:fake-device]",
                  L"MTP drive info: volume label should use the device root segment.");
    state.Require(drive.fileSystem && std::wstring_view(drive.fileSystem) == L"MTP", L"MTP drive info: file-system label mismatch.");
    state.Require((drive.flags & (DRIVE_INFO_FLAG_HAS_TOTAL_BYTES | DRIVE_INFO_FLAG_HAS_FREE_BYTES | DRIVE_INFO_FLAG_HAS_USED_BYTES)) == 0,
                  L"MTP drive info: v1 should not advertise synthetic capacity numbers.");
    state.Require(driveInfoService->GetDriveInfo(L"/", nullptr) == E_POINTER, L"MTP drive info: null DriveInfo pointer should return E_POINTER.");

    const NavigationMenuItem* menuItems = nullptr;
    unsigned int menuCount              = 0;
    hr                                  = driveInfoService->GetDriveMenuItems(L"/", &menuItems, &menuCount);
    state.Require(SUCCEEDED(hr) && menuItems != nullptr && menuCount == 2u,
                  std::format(L"MTP drive info: GetDriveMenuItems expected two entries. hr=0x{:08X} count={}", static_cast<unsigned long>(hr), menuCount));
    if (FAILED(hr) || menuItems == nullptr || menuCount != 2u)
    {
        return false;
    }

    state.Require(menuItems[0].commandId == DRIVE_INFO_COMMAND_PROPERTIES, L"MTP drive info: first menu command should be Properties.");
    state.Require((menuItems[0].flags & NAV_MENU_ITEM_FLAG_DISABLED) != 0, L"MTP drive info: Properties command should be disabled.");
    state.Require(menuItems[0].label != nullptr && menuItems[0].label[0] != L'\0', L"MTP drive info: Properties label is empty.");
    state.Require(menuItems[1].commandId == DRIVE_INFO_COMMAND_CLEANUP, L"MTP drive info: second menu command should be Disconnect.");
    state.Require((menuItems[1].flags & NAV_MENU_ITEM_FLAG_DISABLED) == 0, L"MTP drive info: Disconnect command should be enabled.");
    state.Require(menuItems[1].label != nullptr && menuItems[1].label[0] != L'\0', L"MTP drive info: Disconnect label is empty.");

    const HRESULT propertiesHr = driveInfoService->ExecuteDriveMenuCommand(DRIVE_INFO_COMMAND_PROPERTIES, L"/");
    state.Require(propertiesHr == HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED),
                  std::format(L"MTP drive info: Properties command expected ERROR_NOT_SUPPORTED, got 0x{:08X}.", static_cast<unsigned long>(propertiesHr)));

    const HRESULT disconnectHr = driveInfoService->ExecuteDriveMenuCommand(DRIVE_INFO_COMMAND_CLEANUP, L"/");
    state.Require(SUCCEEDED(disconnectHr), std::format(L"MTP drive info: Disconnect command failed. hr=0x{:08X}", static_cast<unsigned long>(disconnectHr)));

    wil::com_ptr<IFilesInformation> disconnectedInfo;
    const HRESULT enumAfterDisconnectHr = created.fileSystem->ReadDirectoryInfo(L"/", disconnectedInfo.put());
    state.Require(enumAfterDisconnectHr == kDeviceDisconnected && ! disconnectedInfo,
                  std::format(L"MTP drive info: enumeration after Disconnect expected device-gone/null, got hr=0x{:08X}.",
                              static_cast<unsigned long>(enumAfterDisconnectHr)));

    const char* capabilities = nullptr;
    hr                       = created.fileSystem->GetCapabilities(&capabilities);
    state.Require(SUCCEEDED(hr) && capabilities,
                  std::format(L"MTP drive info: capabilities after Disconnect failed. hr=0x{:08X}", static_cast<unsigned long>(hr)));
    if (SUCCEEDED(hr) && capabilities)
    {
        const std::string_view capabilityText(capabilities);
        state.Require(capabilityText.find(R"json("read": false)json") != std::string_view::npos,
                      L"MTP drive info: capabilities should report read=false after Disconnect.");
        state.Require(capabilityText.find(R"json("properties": false)json") != std::string_view::npos,
                      L"MTP drive info: capabilities should report properties=false after Disconnect.");
        state.Require(capabilityText.find(R"json("write": false)json") != std::string_view::npos,
                      L"MTP drive info: capabilities should report write=false after Disconnect.");
    }

    const NavigationMenuItem* afterDisconnectItems = nullptr;
    unsigned int afterDisconnectCount              = 0;
    hr                                             = driveInfoService->GetDriveMenuItems(L"/", &afterDisconnectItems, &afterDisconnectCount);
    state.Require(SUCCEEDED(hr) && afterDisconnectItems != nullptr && afterDisconnectCount == 2u,
                  L"MTP drive info: drive menu should remain available after Disconnect.");

    return state.failure.empty();
});

SelfTest::RunCase(options,
                  suite,
                  L"mtp_hung_device_times_out",
                  [&](SelfTest::CaseState& state) noexcept
{
    CreatedFileSystemInstance created;
    const HRESULT createHr =
        TryCreateFakeMtpFileSystemInstance(R"json({"readFileDelayMs":2000})json", R"json({"readOnly":true,"commandTimeoutMs":50})json", L"/", created);
    if (createHr == HRESULT_FROM_WIN32(ERROR_PROC_NOT_FOUND))
    {
        return state.Skip(L"MTP fake backend export is available only in debug plugin builds.");
    }
    state.Require(SUCCEEDED(createHr) && created.fileSystem && created.module,
                  std::format(L"MTP hung watchdog: create selftest instance failed. hr=0x{:08X}", static_cast<unsigned long>(createHr)));
    if (FAILED(createHr) || ! created.fileSystem || ! created.module)
    {
        return false;
    }

#pragma warning(push)
#pragma warning(disable : 4191) // unsafe conversion from FARPROC
    const auto shutdown     = reinterpret_cast<PluginShutdownFunc>(GetProcAddress(created.module.get(), "RedSalamanderPluginShutdown"));
    const auto canUnloadNow = reinterpret_cast<PluginCanUnloadNowFunc>(GetProcAddress(created.module.get(), "RedSalamanderPluginCanUnloadNow"));
#pragma warning(pop)
    state.Require(shutdown != nullptr, L"MTP hung watchdog: RedSalamanderPluginShutdown export missing.");
    state.Require(canUnloadNow != nullptr, L"MTP hung watchdog: RedSalamanderPluginCanUnloadNow export missing.");
    if (shutdown == nullptr || canUnloadNow == nullptr)
    {
        return false;
    }

    wil::com_ptr<IFileSystemIO> io;
    state.Require(CreateFileSystemIo(created.fileSystem, io), L"MTP hung watchdog: missing IFileSystemIO.");
    if (! io)
    {
        return false;
    }

    constexpr std::wstring_view kPhotoPath = L"/Fake Phone [devpuid:fake-device]/Internal Storage/DCIM/Camera/photo001.txt";
    constexpr HRESULT kDeviceGone          = HRESULT_FROM_WIN32(ERROR_DEVICE_NOT_CONNECTED);
    wil::com_ptr<IFileReader> reader;
    const HRESULT readerHr = io->CreateFileReader(kPhotoPath.data(), reader.put());
    state.Require(SUCCEEDED(readerHr) && reader,
                  std::format(L"MTP hung watchdog: CreateFileReader failed before the delayed Read. hr=0x{:08X}", static_cast<unsigned long>(readerHr)));
    if (FAILED(readerHr) || ! reader)
    {
        return false;
    }

    char byte                 = 0;
    unsigned long bytesRead   = 0;
    const auto start          = std::chrono::steady_clock::now();
    const HRESULT readHr      = reader->Read(&byte, 1u, &bytesRead);
    const auto elapsedMs      = std::chrono::duration_cast<std::chrono::milliseconds>(std::chrono::steady_clock::now() - start).count();
    state.Require(readHr == kDeviceGone && bytesRead == 0u,
                  std::format(L"MTP hung watchdog: delayed Read expected device-gone/0 bytes, got hr=0x{:08X} bytes={}.",
                              static_cast<unsigned long>(readHr),
                              bytesRead));
    state.Require(elapsedMs < 1000, std::format(L"MTP hung watchdog: delayed Read returned after {} ms; expected watchdog timeout.", elapsedMs));

    wil::com_ptr<IFilesInformation> disconnectedInfo;
    const HRESULT enumAfterTimeoutHr = created.fileSystem->ReadDirectoryInfo(L"/", disconnectedInfo.put());
    state.Require(
        enumAfterTimeoutHr == kDeviceGone && ! disconnectedInfo,
        std::format(L"MTP hung watchdog: instance should be device-gone after timeout, got hr=0x{:08X}.", static_cast<unsigned long>(enumAfterTimeoutHr)));

    wil::com_ptr<IFileSystemInitialize> initializer;
    const HRESULT initializerHr = created.fileSystem->QueryInterface(__uuidof(IFileSystemInitialize), initializer.put_void());
    state.Require(SUCCEEDED(initializerHr) && initializer, L"MTP hung watchdog: missing IFileSystemInitialize.");
    const HRESULT reinitializeHr = initializer ? initializer->Initialize(L"/", R"json({"readOnly":true,"commandTimeoutMs":50})json") : initializerHr;
    state.Require(SUCCEEDED(reinitializeHr),
                  std::format(L"MTP hung watchdog: Initialize did not recreate the backend worker. hr=0x{:08X}.",
                              static_cast<unsigned long>(reinitializeHr)));
    unsigned long attributes = 0;
    const HRESULT recoveredHr = io->GetAttributes(kPhotoPath.data(), &attributes);
    state.Require(SUCCEEDED(recoveredHr) && (attributes & FILE_ATTRIBUTE_DIRECTORY) == 0,
                  std::format(L"MTP hung watchdog: command after Initialize did not recover. hr=0x{:08X}.",
                              static_cast<unsigned long>(recoveredHr)));

    bytesRead                 = 0;
    const auto staleStart     = std::chrono::steady_clock::now();
    const HRESULT staleReadHr = reader->Read(&byte, 1u, &bytesRead);
    const auto staleElapsedMs =
        std::chrono::duration_cast<std::chrono::milliseconds>(std::chrono::steady_clock::now() - staleStart).count();
    state.Require(staleReadHr == kDeviceGone && bytesRead == 0u,
                  std::format(L"MTP hung watchdog: stale pre-timeout reader expected device-gone/0 bytes, got hr=0x{:08X} bytes={}.",
                              static_cast<unsigned long>(staleReadHr),
                              bytesRead));
    state.Require(staleElapsedMs < 250,
                  std::format(L"MTP hung watchdog: stale reader took {} ms; generation rejection should be immediate.", staleElapsedMs));

    reader.reset();
    io.reset();
    created.fileSystem.reset();

    shutdown();
    state.Require(canUnloadNow() == FALSE, L"MTP hung watchdog: module reported unloadable while timed-out command was still quarantined.");

    bool becameUnloadable = false;
    for (uint32_t attempt = 0; attempt < 80u; ++attempt)
    {
        std::this_thread::sleep_for(std::chrono::milliseconds(50));
        shutdown();
        if (canUnloadNow() == TRUE)
        {
            becameUnloadable = true;
            break;
        }
    }
    state.Require(becameUnloadable, L"MTP hung watchdog: module did not become unloadable after delayed fake backend returned.");

    return state.failure.empty();
});

SelfTest::RunCase(options,
                  suite,
                  L"mtp_watchdog_requests_backend_cancel",
                  [&](SelfTest::CaseState& state) noexcept
{
    CreatedFileSystemInstance created;
    const HRESULT createHr = TryCreateFakeMtpFileSystemInstance(
        R"json({"operationDelayMs":5000,"cancelUnblocksDelay":true})json", R"json({"readOnly":true,"commandTimeoutMs":50})json", L"/", created);
    if (createHr == HRESULT_FROM_WIN32(ERROR_PROC_NOT_FOUND))
    {
        return state.Skip(L"MTP fake backend export is available only in debug plugin builds.");
    }
    state.Require(SUCCEEDED(createHr) && created.fileSystem && created.module,
                  std::format(L"MTP watchdog cancel: create selftest instance failed. hr=0x{:08X}", static_cast<unsigned long>(createHr)));
    if (FAILED(createHr) || ! created.fileSystem || ! created.module)
    {
        return false;
    }

#pragma warning(push)
#pragma warning(disable : 4191) // unsafe conversion from FARPROC
    const auto shutdown     = reinterpret_cast<PluginShutdownFunc>(GetProcAddress(created.module.get(), "RedSalamanderPluginShutdown"));
    const auto canUnloadNow = reinterpret_cast<PluginCanUnloadNowFunc>(GetProcAddress(created.module.get(), "RedSalamanderPluginCanUnloadNow"));
#pragma warning(pop)
    state.Require(shutdown != nullptr, L"MTP watchdog cancel: RedSalamanderPluginShutdown export missing.");
    state.Require(canUnloadNow != nullptr, L"MTP watchdog cancel: RedSalamanderPluginCanUnloadNow export missing.");
    if (shutdown == nullptr || canUnloadNow == nullptr)
    {
        return false;
    }

    wil::com_ptr<IFileSystemIO> io;
    state.Require(CreateFileSystemIo(created.fileSystem, io), L"MTP watchdog cancel: missing IFileSystemIO.");
    if (! io)
    {
        return false;
    }

    constexpr std::wstring_view kPhotoPath = L"/Fake Phone [devpuid:fake-device]/Internal Storage/DCIM/Camera/photo001.txt";
    constexpr HRESULT kDeviceGone          = HRESULT_FROM_WIN32(ERROR_DEVICE_NOT_CONNECTED);
    wil::com_ptr<IFileReader> reader;
    const auto start     = std::chrono::steady_clock::now();
    const HRESULT readHr = io->CreateFileReader(kPhotoPath.data(), reader.put());
    const auto elapsedMs = std::chrono::duration_cast<std::chrono::milliseconds>(std::chrono::steady_clock::now() - start).count();
    state.Require(readHr == kDeviceGone && ! reader,
                  std::format(L"MTP watchdog cancel: delayed read expected device-gone/null, got hr=0x{:08X}.", static_cast<unsigned long>(readHr)));
    state.Require(elapsedMs < 1000, std::format(L"MTP watchdog cancel: delayed read returned after {} ms; expected watchdog timeout.", elapsedMs));

    io.reset();
    created.fileSystem.reset();

    bool becameUnloadable   = false;
    const auto cleanupStart = std::chrono::steady_clock::now();
    for (uint32_t attempt = 0; attempt < 60u; ++attempt)
    {
        std::this_thread::sleep_for(std::chrono::milliseconds(50));
        shutdown();
        if (canUnloadNow() == TRUE)
        {
            becameUnloadable = true;
            break;
        }
    }
    const auto cleanupElapsedMs = std::chrono::duration_cast<std::chrono::milliseconds>(std::chrono::steady_clock::now() - cleanupStart).count();
    state.Require(becameUnloadable, L"MTP watchdog cancel: backend cancel request did not unblock the delayed fake backend before the original delay.");
    state.Require(cleanupElapsedMs < 4000,
                  std::format(L"MTP watchdog cancel: cleanup took {} ms; expected RequestCancel to shorten the 5000 ms backend delay.", cleanupElapsedMs));

    return state.failure.empty();
});

SelfTest::RunCase(options,
                  suite,
                  L"mtp_runtime_refresh_defers_when_worker_quarantined",
                  [&](SelfTest::CaseState& state) noexcept
{
    FileSystemPluginManager& pluginManager                          = FileSystemPluginManager::GetInstance();
    const Common::Settings::PluginsSettings originalPluginsSettings = g_settings.plugins;
    auto restoreSettings                                            = wil::scope_exit([&]() noexcept
    {
        g_settings.plugins = originalPluginsSettings;
        static_cast<void>(pluginManager.Refresh(g_settings));
    });

    const FileSystemPluginManager::PluginEntry* initialEntry = FindFileSystemPluginById(kBuiltinMtpFileSystemId);
    state.Require(initialEntry != nullptr && initialEntry->loadable && ! initialEntry->unloadDeferred && ! initialEntry->path.empty(),
                  L"MTP runtime refresh: initial manager entry was not a loadable MTP plugin.");
    if (initialEntry == nullptr || initialEntry->path.empty())
    {
        return false;
    }

    const std::filesystem::path mtpPluginPath = initialEntry->path;
    CreatedFileSystemInstance created;
    const HRESULT createHr =
        TryCreateFakeMtpFileSystemInstance(R"json({"readFileDelayMs":5000})json", R"json({"readOnly":true,"commandTimeoutMs":50})json", L"/", created);
    if (createHr == HRESULT_FROM_WIN32(ERROR_PROC_NOT_FOUND))
    {
        return state.Skip(L"MTP fake backend export is available only in debug plugin builds.");
    }
    state.Require(SUCCEEDED(createHr) && created.fileSystem && created.module,
                  std::format(L"MTP runtime refresh: create selftest instance failed. hr=0x{:08X}", static_cast<unsigned long>(createHr)));
    if (FAILED(createHr) || ! created.fileSystem || ! created.module)
    {
        return false;
    }

#pragma warning(push)
#pragma warning(disable : 4191) // unsafe conversion from FARPROC
    const auto shutdown     = reinterpret_cast<PluginShutdownFunc>(GetProcAddress(created.module.get(), "RedSalamanderPluginShutdown"));
    const auto canUnloadNow = reinterpret_cast<PluginCanUnloadNowFunc>(GetProcAddress(created.module.get(), "RedSalamanderPluginCanUnloadNow"));
#pragma warning(pop)
    state.Require(shutdown != nullptr, L"MTP runtime refresh: RedSalamanderPluginShutdown export missing.");
    state.Require(canUnloadNow != nullptr, L"MTP runtime refresh: RedSalamanderPluginCanUnloadNow export missing.");
    if (shutdown == nullptr || canUnloadNow == nullptr)
    {
        return false;
    }

    wil::com_ptr<IFileSystemIO> io;
    state.Require(CreateFileSystemIo(created.fileSystem, io), L"MTP runtime refresh: missing IFileSystemIO.");
    if (! io)
    {
        return false;
    }

    constexpr std::wstring_view kPhotoPath = L"/Fake Phone [devpuid:fake-device]/Internal Storage/DCIM/Camera/photo001.txt";
    wil::com_ptr<IFileReader> reader;
    const HRESULT readerHr = io->CreateFileReader(kPhotoPath.data(), reader.put());
    state.Require(SUCCEEDED(readerHr) && reader,
                  std::format(L"MTP runtime refresh: CreateFileReader failed before the delayed Read. hr=0x{:08X}", static_cast<unsigned long>(readerHr)));
    if (FAILED(readerHr) || ! reader)
    {
        return false;
    }

    char byte               = 0;
    unsigned long bytesRead = 0;
    const HRESULT readHr    = reader->Read(&byte, 1u, &bytesRead);
    state.Require(readHr == HRESULT_FROM_WIN32(ERROR_DEVICE_NOT_CONNECTED) && bytesRead == 0u,
                  std::format(L"MTP runtime refresh: delayed Read expected device-gone/0 bytes, got hr=0x{:08X} bytes={}.",
                              static_cast<unsigned long>(readHr),
                              bytesRead));

    reader.reset();
    io.reset();
    created.fileSystem.reset();

    shutdown();
    state.Require(canUnloadNow() == FALSE, L"MTP runtime refresh: test setup did not leave a quarantined MTP worker.");

    const HRESULT deferredRefreshHr = pluginManager.Refresh(g_settings);
    state.Require(SUCCEEDED(deferredRefreshHr),
                  std::format(L"MTP runtime refresh: manager refresh while quarantined failed. hr=0x{:08X}", static_cast<unsigned long>(deferredRefreshHr)));
    if (FAILED(deferredRefreshHr))
    {
        return false;
    }

    const FileSystemPluginManager::PluginEntry* deferredEntry = FindFileSystemPluginById(kBuiltinMtpFileSystemId);
    state.Require(deferredEntry != nullptr, L"MTP runtime refresh: deferred MTP placeholder missing after refresh.");
    if (deferredEntry != nullptr)
    {
        state.Require(deferredEntry->unloadDeferred, L"MTP runtime refresh: MTP entry should be marked unload-deferred.");
        state.Require(! deferredEntry->loadable, L"MTP runtime refresh: deferred MTP entry should not be loadable.");
        state.Require(deferredEntry->path == mtpPluginPath, L"MTP runtime refresh: deferred MTP entry path changed.");
        state.Require(deferredEntry->module.get() == nullptr, L"MTP runtime refresh: public deferred placeholder should not expose a reload module.");
    }

    bool becameUnloadable = false;
    for (uint32_t attempt = 0; attempt < 140u; ++attempt)
    {
        std::this_thread::sleep_for(std::chrono::milliseconds(50));
        shutdown();
        if (canUnloadNow() == TRUE)
        {
            becameUnloadable = true;
            break;
        }
    }
    state.Require(becameUnloadable, L"MTP runtime refresh: quarantined worker did not complete.");

    created.module.reset();

    const HRESULT restoredRefreshHr = pluginManager.Refresh(g_settings);
    state.Require(
        SUCCEEDED(restoredRefreshHr),
        std::format(L"MTP runtime refresh: manager refresh after worker completion failed. hr=0x{:08X}", static_cast<unsigned long>(restoredRefreshHr)));
    if (FAILED(restoredRefreshHr))
    {
        return false;
    }

    const FileSystemPluginManager::PluginEntry* restoredEntry = FindFileSystemPluginById(kBuiltinMtpFileSystemId);
    state.Require(restoredEntry != nullptr && restoredEntry->loadable && ! restoredEntry->unloadDeferred && restoredEntry->path == mtpPluginPath,
                  L"MTP runtime refresh: MTP plugin was not restored as loadable after deferred unload completed.");

    return state.failure.empty();
});

SelfTest::RunCase(options,
                  suite,
                  L"mtp_mutating_create_directory_times_out",
                  [&](SelfTest::CaseState& state) noexcept
{
    CreatedFileSystemInstance created;
    const HRESULT createHr =
        TryCreateFakeMtpFileSystemInstance(R"json({"operationDelayMs":2000})json", R"json({"readOnly":false,"commandTimeoutMs":50})json", L"/", created);
    if (createHr == HRESULT_FROM_WIN32(ERROR_PROC_NOT_FOUND))
    {
        return state.Skip(L"MTP fake backend export is available only in debug plugin builds.");
    }
    state.Require(SUCCEEDED(createHr) && created.fileSystem && created.module,
                  std::format(L"MTP mutation watchdog: create selftest instance failed. hr=0x{:08X}", static_cast<unsigned long>(createHr)));
    if (FAILED(createHr) || ! created.fileSystem || ! created.module)
    {
        return false;
    }

#pragma warning(push)
#pragma warning(disable : 4191) // unsafe conversion from FARPROC
    const auto shutdown     = reinterpret_cast<PluginShutdownFunc>(GetProcAddress(created.module.get(), "RedSalamanderPluginShutdown"));
    const auto canUnloadNow = reinterpret_cast<PluginCanUnloadNowFunc>(GetProcAddress(created.module.get(), "RedSalamanderPluginCanUnloadNow"));
#pragma warning(pop)
    state.Require(shutdown != nullptr, L"MTP mutation watchdog: RedSalamanderPluginShutdown export missing.");
    state.Require(canUnloadNow != nullptr, L"MTP mutation watchdog: RedSalamanderPluginCanUnloadNow export missing.");
    if (shutdown == nullptr || canUnloadNow == nullptr)
    {
        return false;
    }

    wil::com_ptr<IFileSystemDirectoryOperations> dirOps;
    state.Require(CreateFileSystemDirectoryOperations(created.fileSystem, dirOps), L"MTP mutation watchdog: missing directory operations.");
    if (! dirOps)
    {
        return false;
    }

    const std::wstring targetPath = L"/Fake Phone [devpuid:fake-device]/Internal Storage/DCIM/Camera/watchdog-created-" + MakeGuidText();
    constexpr HRESULT kDeviceGone = HRESULT_FROM_WIN32(ERROR_DEVICE_NOT_CONNECTED);
    const auto start              = std::chrono::steady_clock::now();
    const HRESULT mkdirHr         = dirOps->CreateDirectory(targetPath.c_str());
    const auto elapsedMs          = std::chrono::duration_cast<std::chrono::milliseconds>(std::chrono::steady_clock::now() - start).count();
    state.Require(mkdirHr == kDeviceGone,
                  std::format(L"MTP mutation watchdog: delayed CreateDirectory expected device-gone, got hr=0x{:08X}.", static_cast<unsigned long>(mkdirHr)));
    state.Require(elapsedMs < 1000, std::format(L"MTP mutation watchdog: delayed CreateDirectory returned after {} ms; expected watchdog timeout.", elapsedMs));

    wil::com_ptr<IFilesInformation> disconnectedInfo;
    const HRESULT enumAfterTimeoutHr = created.fileSystem->ReadDirectoryInfo(L"/", disconnectedInfo.put());
    state.Require(enumAfterTimeoutHr == kDeviceGone && ! disconnectedInfo,
                  std::format(L"MTP mutation watchdog: instance should be device-gone after mutation timeout, got hr=0x{:08X}.",
                              static_cast<unsigned long>(enumAfterTimeoutHr)));

    dirOps.reset();
    created.fileSystem.reset();

    shutdown();
    state.Require(canUnloadNow() == FALSE, L"MTP mutation watchdog: module reported unloadable while mutating command was quarantined.");

    bool becameUnloadable = false;
    for (uint32_t attempt = 0; attempt < 80u; ++attempt)
    {
        std::this_thread::sleep_for(std::chrono::milliseconds(50));
        shutdown();
        if (canUnloadNow() == TRUE)
        {
            becameUnloadable = true;
            break;
        }
    }
    state.Require(becameUnloadable, L"MTP mutation watchdog: module did not become unloadable after delayed mutating command returned.");

    return state.failure.empty();
});

SelfTest::RunCase(options,
                  suite,
                  L"mtp_mutating_item_commands_time_out",
                  [&](SelfTest::CaseState& state) noexcept
{
    constexpr std::wstring_view kPhotoPath = L"/Fake Phone [devpuid:fake-device]/Internal Storage/DCIM/Camera/photo001.txt";

    const auto runTimedOutMutation = [&](std::wstring_view label, const auto& invoke) noexcept -> bool
    {
        CreatedFileSystemInstance created;
        const HRESULT createHr =
            TryCreateFakeMtpFileSystemInstance(R"json({"operationDelayMs":2000})json", R"json({"readOnly":false,"commandTimeoutMs":50})json", L"/", created);
        if (createHr == HRESULT_FROM_WIN32(ERROR_PROC_NOT_FOUND))
        {
            return state.Skip(L"MTP fake backend export is available only in debug plugin builds.");
        }
        state.Require(
            SUCCEEDED(createHr) && created.fileSystem && created.module,
            std::format(L"MTP item mutation watchdog ({}): create selftest instance failed. hr=0x{:08X}", label, static_cast<unsigned long>(createHr)));
        if (FAILED(createHr) || ! created.fileSystem || ! created.module)
        {
            return false;
        }

#pragma warning(push)
#pragma warning(disable : 4191) // unsafe conversion from FARPROC
        const auto shutdown     = reinterpret_cast<PluginShutdownFunc>(GetProcAddress(created.module.get(), "RedSalamanderPluginShutdown"));
        const auto canUnloadNow = reinterpret_cast<PluginCanUnloadNowFunc>(GetProcAddress(created.module.get(), "RedSalamanderPluginCanUnloadNow"));
#pragma warning(pop)
        state.Require(shutdown != nullptr, std::format(L"MTP item mutation watchdog ({}): RedSalamanderPluginShutdown export missing.", label));
        state.Require(canUnloadNow != nullptr, std::format(L"MTP item mutation watchdog ({}): RedSalamanderPluginCanUnloadNow export missing.", label));
        if (shutdown == nullptr || canUnloadNow == nullptr)
        {
            return false;
        }

        const auto start         = std::chrono::steady_clock::now();
        const HRESULT mutationHr = invoke(created.fileSystem.get());
        const auto elapsedMs     = std::chrono::duration_cast<std::chrono::milliseconds>(std::chrono::steady_clock::now() - start).count();
        state.Require(mutationHr == HRESULT_FROM_WIN32(ERROR_DEVICE_NOT_CONNECTED),
                      std::format(L"MTP item mutation watchdog ({}): expected device-gone, got hr=0x{:08X}.", label, static_cast<unsigned long>(mutationHr)));
        state.Require(elapsedMs < 1000,
                      std::format(L"MTP item mutation watchdog ({}): command returned after {} ms; expected watchdog timeout.", label, elapsedMs));

        wil::com_ptr<IFilesInformation> disconnectedInfo;
        const HRESULT enumAfterTimeoutHr = created.fileSystem->ReadDirectoryInfo(L"/", disconnectedInfo.put());
        state.Require(enumAfterTimeoutHr == HRESULT_FROM_WIN32(ERROR_DEVICE_NOT_CONNECTED) && ! disconnectedInfo,
                      std::format(L"MTP item mutation watchdog ({}): instance should be device-gone after timeout, got hr=0x{:08X}.",
                                  label,
                                  static_cast<unsigned long>(enumAfterTimeoutHr)));

        created.fileSystem.reset();

        shutdown();
        state.Require(canUnloadNow() == FALSE,
                      std::format(L"MTP item mutation watchdog ({}): module reported unloadable while command was quarantined.", label));

        bool becameUnloadable = false;
        for (uint32_t attempt = 0; attempt < 80u; ++attempt)
        {
            std::this_thread::sleep_for(std::chrono::milliseconds(50));
            shutdown();
            if (canUnloadNow() == TRUE)
            {
                becameUnloadable = true;
                break;
            }
        }
        state.Require(becameUnloadable,
                      std::format(L"MTP item mutation watchdog ({}): module did not become unloadable after delayed command returned.", label));

        return state.failure.empty();
    };

    std::wstring guid = MakeGuidText();
    state.Require(! guid.empty(), L"MTP item mutation watchdog: failed to generate copy destination name.");
    if (guid.empty())
    {
        return false;
    }
    const std::wstring copyDestination = L"/Fake Phone [devpuid:fake-device]/Internal Storage/DCIM/Camera/watchdog-copy-" + guid + L".txt";
    if (! runTimedOutMutation(L"CopyItem", [&](IFileSystem* fileSystem) noexcept {
        return fileSystem->CopyItem(kPhotoPath.data(), copyDestination.c_str(), FILESYSTEM_FLAG_NONE, nullptr, nullptr, nullptr);
    }))
    {
        return false;
    }

    guid = MakeGuidText();
    state.Require(! guid.empty(), L"MTP item mutation watchdog: failed to generate move destination name.");
    if (guid.empty())
    {
        return false;
    }
    const std::wstring moveDestination = L"/Fake Phone [devpuid:fake-device]/Internal Storage/DCIM/Camera/watchdog-move-" + guid + L".txt";
    if (! runTimedOutMutation(L"MoveItem", [&](IFileSystem* fileSystem) noexcept {
        return fileSystem->MoveItem(kPhotoPath.data(), moveDestination.c_str(), FILESYSTEM_FLAG_NONE, nullptr, nullptr, nullptr);
    }))
    {
        return false;
    }

    guid = MakeGuidText();
    state.Require(! guid.empty(), L"MTP item mutation watchdog: failed to generate rename destination name.");
    if (guid.empty())
    {
        return false;
    }
    const std::wstring renameDestination = L"/Fake Phone [devpuid:fake-device]/Internal Storage/DCIM/Camera/watchdog-rename-" + guid + L".txt";
    if (! runTimedOutMutation(L"RenameItem", [&](IFileSystem* fileSystem) noexcept {
        return fileSystem->RenameItem(kPhotoPath.data(), renameDestination.c_str(), FILESYSTEM_FLAG_NONE, nullptr, nullptr, nullptr);
    }))
    {
        return false;
    }

    if (! runTimedOutMutation(L"DeleteItem", [&](IFileSystem* fileSystem) noexcept {
        return fileSystem->DeleteItem(kPhotoPath.data(), FILESYSTEM_FLAG_NONE, nullptr, nullptr, nullptr);
    }))
    {
        return false;
    }

    return state.failure.empty();
});

SelfTest::RunCase(options,
                  suite,
                  L"mtp_writer_commit_times_out",
                  [&](SelfTest::CaseState& state) noexcept
{
    CreatedFileSystemInstance created;
    const HRESULT createHr =
        TryCreateFakeMtpFileSystemInstance(R"json({"operationDelayMs":2000})json", R"json({"readOnly":false,"commandTimeoutMs":50})json", L"/", created);
    if (createHr == HRESULT_FROM_WIN32(ERROR_PROC_NOT_FOUND))
    {
        return state.Skip(L"MTP fake backend export is available only in debug plugin builds.");
    }
    state.Require(SUCCEEDED(createHr) && created.fileSystem && created.module,
                  std::format(L"MTP writer watchdog: create selftest instance failed. hr=0x{:08X}", static_cast<unsigned long>(createHr)));
    if (FAILED(createHr) || ! created.fileSystem || ! created.module)
    {
        return false;
    }

#pragma warning(push)
#pragma warning(disable : 4191) // unsafe conversion from FARPROC
    const auto shutdown     = reinterpret_cast<PluginShutdownFunc>(GetProcAddress(created.module.get(), "RedSalamanderPluginShutdown"));
    const auto canUnloadNow = reinterpret_cast<PluginCanUnloadNowFunc>(GetProcAddress(created.module.get(), "RedSalamanderPluginCanUnloadNow"));
#pragma warning(pop)
    state.Require(shutdown != nullptr, L"MTP writer watchdog: RedSalamanderPluginShutdown export missing.");
    state.Require(canUnloadNow != nullptr, L"MTP writer watchdog: RedSalamanderPluginCanUnloadNow export missing.");
    if (shutdown == nullptr || canUnloadNow == nullptr)
    {
        return false;
    }

    wil::com_ptr<IFileSystemIO> io;
    state.Require(CreateFileSystemIo(created.fileSystem, io), L"MTP writer watchdog: missing IFileSystemIO.");
    if (! io)
    {
        return false;
    }

    const std::wstring guid = MakeGuidText();
    state.Require(! guid.empty(), L"MTP writer watchdog: failed to generate unique file name.");
    if (guid.empty())
    {
        return false;
    }

    const std::wstring targetPath = L"/Fake Phone [devpuid:fake-device]/Internal Storage/DCIM/Camera/watchdog-writer-" + guid + L".txt";

    wil::com_ptr<IFileWriter> writer;
    const HRESULT writerHr = io->CreateFileWriter(targetPath.c_str(), FILESYSTEM_FLAG_NONE, writer.put());
    state.Require(SUCCEEDED(writerHr) && writer,
                  std::format(L"MTP writer watchdog: CreateFileWriter failed. hr=0x{:08X}", static_cast<unsigned long>(writerHr)));
    if (FAILED(writerHr) || ! writer)
    {
        return false;
    }

    constexpr std::string_view kPayload = "writer watchdog payload";
    static_assert(kPayload.size() <= static_cast<size_t>((std::numeric_limits<unsigned long>::max)()));
    unsigned long written = 0;
    const HRESULT writeHr = writer->Write(kPayload.data(), static_cast<unsigned long>(kPayload.size()), &written);
    state.Require(SUCCEEDED(writeHr) && written == static_cast<unsigned long>(kPayload.size()),
                  std::format(L"MTP writer watchdog: Write failed. bytes={} hr=0x{:08X}.", written, static_cast<unsigned long>(writeHr)));
    if (FAILED(writeHr) || written != static_cast<unsigned long>(kPayload.size()))
    {
        return false;
    }

    const auto start       = std::chrono::steady_clock::now();
    const HRESULT commitHr = writer->Commit();
    const auto elapsedMs   = std::chrono::duration_cast<std::chrono::milliseconds>(std::chrono::steady_clock::now() - start).count();
    state.Require(commitHr == HRESULT_FROM_WIN32(ERROR_DEVICE_NOT_CONNECTED),
                  std::format(L"MTP writer watchdog: expected device-gone, got hr=0x{:08X}.", static_cast<unsigned long>(commitHr)));
    state.Require(elapsedMs < 1000, std::format(L"MTP writer watchdog: Commit returned after {} ms; expected watchdog timeout.", elapsedMs));

    wil::com_ptr<IFilesInformation> disconnectedInfo;
    const HRESULT enumAfterTimeoutHr = created.fileSystem->ReadDirectoryInfo(L"/", disconnectedInfo.put());
    state.Require(
        enumAfterTimeoutHr == HRESULT_FROM_WIN32(ERROR_DEVICE_NOT_CONNECTED) && ! disconnectedInfo,
        std::format(L"MTP writer watchdog: instance should be device-gone after timeout, got hr=0x{:08X}.", static_cast<unsigned long>(enumAfterTimeoutHr)));

    writer.reset();
    io.reset();
    created.fileSystem.reset();

    shutdown();
    state.Require(canUnloadNow() == FALSE, L"MTP writer watchdog: module reported unloadable while Commit worker was quarantined.");

    bool becameUnloadable = false;
    for (uint32_t attempt = 0; attempt < 80u; ++attempt)
    {
        std::this_thread::sleep_for(std::chrono::milliseconds(50));
        shutdown();
        if (canUnloadNow() == TRUE)
        {
            becameUnloadable = true;
            break;
        }
    }
    state.Require(becameUnloadable, L"MTP writer watchdog: module did not become unloadable after delayed Commit worker returned.");

    return state.failure.empty();
});

SelfTest::RunCase(options,
                  suite,
                  L"mtp_menu_and_directory_size_time_out",
                  [&](SelfTest::CaseState& state) noexcept
{
    constexpr std::wstring_view kCameraPath = L"/Fake Phone [devpuid:fake-device]/Internal Storage/DCIM/Camera";

    const auto waitForQuarantinedWorker =
        [&](CreatedFileSystemInstance& created, PluginShutdownFunc shutdown, PluginCanUnloadNowFunc canUnloadNow, std::wstring_view label) noexcept -> bool
    {
        created.fileSystem.reset();

        shutdown();
        state.Require(canUnloadNow() == FALSE, std::format(L"{}: module reported unloadable while command was quarantined.", label));

        bool becameUnloadable = false;
        for (uint32_t attempt = 0; attempt < 80u; ++attempt)
        {
            std::this_thread::sleep_for(std::chrono::milliseconds(50));
            shutdown();
            if (canUnloadNow() == TRUE)
            {
                becameUnloadable = true;
                break;
            }
        }
        state.Require(becameUnloadable, std::format(L"{}: module did not become unloadable after delayed command returned.", label));
        return state.failure.empty();
    };

    const auto createDelayedInstance =
        [&](std::wstring_view label, CreatedFileSystemInstance& created, PluginShutdownFunc& shutdown, PluginCanUnloadNowFunc& canUnloadNow) noexcept -> bool
    {
        const HRESULT createHr =
            TryCreateFakeMtpFileSystemInstance(R"json({"operationDelayMs":2000})json", R"json({"readOnly":true,"commandTimeoutMs":50})json", L"/", created);
        if (createHr == HRESULT_FROM_WIN32(ERROR_PROC_NOT_FOUND))
        {
            return state.Skip(L"MTP fake backend export is available only in debug plugin builds.");
        }
        state.Require(SUCCEEDED(createHr) && created.fileSystem && created.module,
                      std::format(L"{}: create selftest instance failed. hr=0x{:08X}", label, static_cast<unsigned long>(createHr)));
        if (FAILED(createHr) || ! created.fileSystem || ! created.module)
        {
            return false;
        }

#pragma warning(push)
#pragma warning(disable : 4191) // unsafe conversion from FARPROC
        shutdown     = reinterpret_cast<PluginShutdownFunc>(GetProcAddress(created.module.get(), "RedSalamanderPluginShutdown"));
        canUnloadNow = reinterpret_cast<PluginCanUnloadNowFunc>(GetProcAddress(created.module.get(), "RedSalamanderPluginCanUnloadNow"));
#pragma warning(pop)
        state.Require(shutdown != nullptr, std::format(L"{}: RedSalamanderPluginShutdown export missing.", label));
        state.Require(canUnloadNow != nullptr, std::format(L"{}: RedSalamanderPluginCanUnloadNow export missing.", label));
        return shutdown != nullptr && canUnloadNow != nullptr && state.failure.empty();
    };

    {
        CreatedFileSystemInstance created;
        PluginShutdownFunc shutdown         = nullptr;
        PluginCanUnloadNowFunc canUnloadNow = nullptr;
        if (! createDelayedInstance(L"MTP menu watchdog", created, shutdown, canUnloadNow))
        {
            return false;
        }

        wil::com_ptr<INavigationMenu> navigationMenu;
        const HRESULT qiHr = created.fileSystem->QueryInterface(__uuidof(INavigationMenu), navigationMenu.put_void());
        state.Require(SUCCEEDED(qiHr) && navigationMenu,
                      std::format(L"MTP menu watchdog: missing INavigationMenu. hr=0x{:08X}", static_cast<unsigned long>(qiHr)));
        if (FAILED(qiHr) || ! navigationMenu)
        {
            return false;
        }

        const NavigationMenuItem* items = nullptr;
        unsigned int count              = 0;
        const auto start                = std::chrono::steady_clock::now();
        const HRESULT menuHr            = navigationMenu->GetMenuItems(&items, &count);
        const auto elapsedMs            = std::chrono::duration_cast<std::chrono::milliseconds>(std::chrono::steady_clock::now() - start).count();
        state.Require(SUCCEEDED(menuHr), std::format(L"MTP menu watchdog: GetMenuItems failed. hr=0x{:08X}", static_cast<unsigned long>(menuHr)));
        state.Require(elapsedMs < 1000, std::format(L"MTP menu watchdog: GetMenuItems returned after {} ms; expected watchdog timeout.", elapsedMs));
        state.Require(items != nullptr && count == 1u && (items[0].flags & NAV_MENU_ITEM_FLAG_DISABLED) != 0,
                      L"MTP menu watchdog: timed-out menu enumeration did not expose the disabled fallback item.");

        wil::com_ptr<IFilesInformation> disconnectedInfo;
        const HRESULT enumAfterTimeoutHr = created.fileSystem->ReadDirectoryInfo(L"/", disconnectedInfo.put());
        state.Require(
            enumAfterTimeoutHr == HRESULT_FROM_WIN32(ERROR_DEVICE_NOT_CONNECTED) && ! disconnectedInfo,
            std::format(L"MTP menu watchdog: instance should be device-gone after timeout, got hr=0x{:08X}.", static_cast<unsigned long>(enumAfterTimeoutHr)));

        navigationMenu.reset();
        if (! waitForQuarantinedWorker(created, shutdown, canUnloadNow, L"MTP menu watchdog"))
        {
            return false;
        }
    }

    {
        CreatedFileSystemInstance created;
        PluginShutdownFunc shutdown         = nullptr;
        PluginCanUnloadNowFunc canUnloadNow = nullptr;
        if (! createDelayedInstance(L"MTP directory-size watchdog", created, shutdown, canUnloadNow))
        {
            return false;
        }

        wil::com_ptr<IFileSystemDirectoryOperations> dirOps;
        state.Require(CreateFileSystemDirectoryOperations(created.fileSystem, dirOps), L"MTP directory-size watchdog: missing directory operations.");
        if (! dirOps)
        {
            return false;
        }

        FileSystemDirectorySizeResult result{};
        result.sizeBytes     = sizeof(FileSystemDirectorySizeResult);
        const auto start     = std::chrono::steady_clock::now();
        const HRESULT sizeHr = dirOps->GetDirectorySize(kCameraPath.data(), FILESYSTEM_FLAG_RECURSIVE, nullptr, nullptr, &result);
        const auto elapsedMs = std::chrono::duration_cast<std::chrono::milliseconds>(std::chrono::steady_clock::now() - start).count();
        state.Require(sizeHr == HRESULT_FROM_WIN32(ERROR_DEVICE_NOT_CONNECTED) && result.status == HRESULT_FROM_WIN32(ERROR_DEVICE_NOT_CONNECTED),
                      std::format(L"MTP directory-size watchdog: expected device-gone, got hr=0x{:08X} status=0x{:08X}.",
                                  static_cast<unsigned long>(sizeHr),
                                  static_cast<unsigned long>(result.status)));
        state.Require(elapsedMs < 1000,
                      std::format(L"MTP directory-size watchdog: GetDirectorySize returned after {} ms; expected watchdog timeout.", elapsedMs));

        wil::com_ptr<IFilesInformation> disconnectedInfo;
        const HRESULT enumAfterTimeoutHr = created.fileSystem->ReadDirectoryInfo(L"/", disconnectedInfo.put());
        state.Require(enumAfterTimeoutHr == HRESULT_FROM_WIN32(ERROR_DEVICE_NOT_CONNECTED) && ! disconnectedInfo,
                      std::format(L"MTP directory-size watchdog: instance should be device-gone after timeout, got hr=0x{:08X}.",
                                  static_cast<unsigned long>(enumAfterTimeoutHr)));

        dirOps.reset();
        if (! waitForQuarantinedWorker(created, shutdown, canUnloadNow, L"MTP directory-size watchdog"))
        {
            return false;
        }
    }

    return state.failure.empty();
});

SelfTest::RunCase(options,
                  suite,
                  L"mtp_reader_seek_contract",
                  [&](SelfTest::CaseState& state) noexcept
{
    CreatedFileSystemInstance created;
    const HRESULT createHr = TryCreateFakeMtpFileSystemInstance("{}", R"json({"readOnly":true})json", L"/", created);
    if (createHr == HRESULT_FROM_WIN32(ERROR_PROC_NOT_FOUND))
    {
        return state.Skip(L"MTP fake backend export is available only in debug plugin builds.");
    }
    state.Require(SUCCEEDED(createHr) && created.fileSystem,
                  std::format(L"MTP reader seek: create selftest instance failed. hr=0x{:08X}", static_cast<unsigned long>(createHr)));
    if (FAILED(createHr) || ! created.fileSystem)
    {
        return false;
    }

    wil::com_ptr<IFileSystemIO> io;
    state.Require(CreateFileSystemIo(created.fileSystem, io), L"MTP reader seek: missing IFileSystemIO.");
    if (! io)
    {
        return false;
    }

    constexpr std::wstring_view kPhotoPath       = L"/Fake Phone [devpuid:fake-device]/Internal Storage/DCIM/Camera/photo001.txt";
    constexpr std::string_view kExpectedContents = "RedSalamander deterministic MTP fixture\r\n";

    wil::com_ptr<IFileReader> reader;
    const HRESULT readerHr = io->CreateFileReader(kPhotoPath.data(), reader.put());
    state.Require(SUCCEEDED(readerHr) && reader, std::format(L"MTP reader seek: CreateFileReader failed. hr=0x{:08X}", static_cast<unsigned long>(readerHr)));
    if (FAILED(readerHr) || ! reader)
    {
        return false;
    }

    uint64_t sizeBytes   = 0;
    const HRESULT sizeHr = reader->GetSize(&sizeBytes);
    state.Require(
        SUCCEEDED(sizeHr) && sizeBytes == kExpectedContents.size(),
        std::format(L"MTP reader seek: expected size {}, got {} hr=0x{:08X}.", kExpectedContents.size(), sizeBytes, static_cast<unsigned long>(sizeHr)));
    state.Require(reader->GetSize(nullptr) == E_POINTER, L"MTP reader seek: GetSize(nullptr) did not return E_POINTER.");

    char first[3]                 = {};
    const unsigned long firstSize = static_cast<unsigned long>(std::size(first));
    unsigned long bytesRead       = 0;
    HRESULT readHr                = reader->Read(first, firstSize, &bytesRead);
    state.Require(SUCCEEDED(readHr) && bytesRead == firstSize && std::string_view(first, bytesRead) == "Red",
                  std::format(L"MTP reader seek: initial read failed. bytes={} hr=0x{:08X}.", bytesRead, static_cast<unsigned long>(readHr)));

    uint64_t newPosition = 0;
    HRESULT seekHr       = reader->Seek(14, FILE_BEGIN, &newPosition);
    state.Require(SUCCEEDED(seekHr) && newPosition == 14u,
                  std::format(L"MTP reader seek: FILE_BEGIN seek expected 14, got {} hr=0x{:08X}.", newPosition, static_cast<unsigned long>(seekHr)));

    char deterministic[13]                = {};
    const unsigned long deterministicSize = static_cast<unsigned long>(std::size(deterministic));
    bytesRead                             = 0;
    readHr                                = reader->Read(deterministic, deterministicSize, &bytesRead);
    state.Require(SUCCEEDED(readHr) && bytesRead == deterministicSize && std::string_view(deterministic, bytesRead) == "deterministic",
                  std::format(L"MTP reader seek: random-position read failed. bytes={} hr=0x{:08X}.", bytesRead, static_cast<unsigned long>(readHr)));

    seekHr = reader->Seek(-1, FILE_CURRENT, &newPosition);
    state.Require(
        SUCCEEDED(seekHr) && newPosition == 26u,
        std::format(L"MTP reader seek: FILE_CURRENT backward seek expected 26, got {} hr=0x{:08X}.", newPosition, static_cast<unsigned long>(seekHr)));

    char tailWord[5]                 = {};
    const unsigned long tailWordSize = static_cast<unsigned long>(std::size(tailWord));
    bytesRead                        = 0;
    readHr                           = reader->Read(tailWord, tailWordSize, &bytesRead);
    state.Require(SUCCEEDED(readHr) && bytesRead == tailWordSize && std::string_view(tailWord, bytesRead) == "c MTP",
                  std::format(L"MTP reader seek: backward seek read failed. bytes={} hr=0x{:08X}.", bytesRead, static_cast<unsigned long>(readHr)));

    seekHr = reader->Seek(-2, FILE_END, &newPosition);
    state.Require(SUCCEEDED(seekHr) && newPosition == kExpectedContents.size() - 2u,
                  std::format(L"MTP reader seek: FILE_END seek expected {}, got {} hr=0x{:08X}.",
                              kExpectedContents.size() - 2u,
                              newPosition,
                              static_cast<unsigned long>(seekHr)));

    char newline[2]                 = {};
    const unsigned long newlineSize = static_cast<unsigned long>(std::size(newline));
    bytesRead                       = 0;
    readHr                          = reader->Read(newline, newlineSize, &bytesRead);
    state.Require(SUCCEEDED(readHr) && bytesRead == newlineSize && std::string_view(newline, bytesRead) == "\r\n",
                  std::format(L"MTP reader seek: tail read failed. bytes={} hr=0x{:08X}.", bytesRead, static_cast<unsigned long>(readHr)));

    seekHr = reader->Seek(999, FILE_BEGIN, &newPosition);
    state.Require(SUCCEEDED(seekHr) && newPosition == 999u,
                  std::format(L"MTP reader seek: beyond EOF seek expected 999, got {} hr=0x{:08X}.", newPosition, static_cast<unsigned long>(seekHr)));
    bytesRead = 42u;
    readHr    = reader->Read(first, firstSize, &bytesRead);
    state.Require(SUCCEEDED(readHr) && bytesRead == 0u,
                  std::format(L"MTP reader seek: beyond EOF read expected 0 bytes, got {} hr=0x{:08X}.", bytesRead, static_cast<unsigned long>(readHr)));

    state.Require(reader->Seek(0, FILE_BEGIN, nullptr) == E_POINTER, L"MTP reader seek: Seek(nullptr) did not return E_POINTER.");
    seekHr = reader->Seek(-1, FILE_BEGIN, &newPosition);
    state.Require(seekHr == HRESULT_FROM_WIN32(ERROR_NEGATIVE_SEEK),
                  std::format(L"MTP reader seek: negative FILE_BEGIN seek expected ERROR_NEGATIVE_SEEK, got 0x{:08X}.", static_cast<unsigned long>(seekHr)));
    seekHr = reader->Seek(0, 0xFFFF'FFFFu, &newPosition);
    state.Require(seekHr == E_INVALIDARG,
                  std::format(L"MTP reader seek: invalid origin expected E_INVALIDARG, got 0x{:08X}.", static_cast<unsigned long>(seekHr)));
    state.Require(reader->Read(nullptr, 1, &bytesRead) == E_POINTER, L"MTP reader seek: Read(nullptr, nonzero) did not return E_POINTER.");
    state.Require(reader->Read(first, 1, nullptr) == E_POINTER, L"MTP reader seek: Read(nullptr bytesRead) did not return E_POINTER.");

    return state.failure.empty();
});

SelfTest::RunCase(options,
                  suite,
                  L"mtp_reader_streams_on_read_not_open",
                  [&](SelfTest::CaseState& state) noexcept
{
    CreatedFileSystemInstance created;
    const HRESULT createHr = TryCreateFakeMtpFileSystemInstance("{}", R"json({"readOnly":true})json", L"/", created);
    if (createHr == HRESULT_FROM_WIN32(ERROR_PROC_NOT_FOUND))
    {
        return state.Skip(L"MTP fake backend export is available only in debug plugin builds.");
    }
    state.Require(SUCCEEDED(createHr) && created.fileSystem,
                  std::format(L"MTP streamed reader: create selftest instance failed. hr=0x{:08X}", static_cast<unsigned long>(createHr)));
    if (FAILED(createHr) || ! created.fileSystem)
    {
        return false;
    }

    wil::com_ptr<IFileSystemIO> io;
    state.Require(CreateFileSystemIo(created.fileSystem, io), L"MTP streamed reader: missing IFileSystemIO.");
    if (! io)
    {
        return false;
    }

    constexpr std::wstring_view kPhotoPath = L"/Fake Phone [devpuid:fake-device]/Internal Storage/DCIM/Camera/photo001.txt";

    const auto readCounter = [&](std::string_view key, uint64_t& value, std::wstring_view label) noexcept -> bool
    {
        const char* properties = nullptr;
        const HRESULT propsHr  = io->GetItemProperties(kPhotoPath.data(), &properties);
        state.Require(SUCCEEDED(propsHr) && properties != nullptr,
                      std::format(L"MTP streamed reader: {} GetItemProperties failed. hr=0x{:08X}", label, static_cast<unsigned long>(propsHr)));
        if (FAILED(propsHr) || properties == nullptr)
        {
            return false;
        }

        const std::optional<uint64_t> parsed = extractJsonUInt(properties, key);
        state.Require(parsed.has_value(), std::format(L"MTP streamed reader: {} missing {} instrumentation.", label, std::wstring(key.begin(), key.end())));
        if (! parsed.has_value())
        {
            return false;
        }

        value = parsed.value();
        return true;
    };

    wil::com_ptr<IFileReader> reader;
    const HRESULT readerHr = io->CreateFileReader(kPhotoPath.data(), reader.put());
    state.Require(SUCCEEDED(readerHr) && reader,
                  std::format(L"MTP streamed reader: CreateFileReader failed. hr=0x{:08X}", static_cast<unsigned long>(readerHr)));
    if (FAILED(readerHr) || ! reader)
    {
        return false;
    }

    uint64_t readFileCalls = 0;
    state.Require(readCounter("readFileCalls", readFileCalls, L"after open"), L"MTP streamed reader: failed to read open counter.");
    state.Require(readFileCalls == 0u,
                  std::format(L"MTP streamed reader: opening the reader materialized the file; readFileCalls={}.", readFileCalls));

    uint64_t sizeBytes = 0;
    const HRESULT sizeHr = reader->GetSize(&sizeBytes);
    state.Require(SUCCEEDED(sizeHr) && sizeBytes == 41u,
                  std::format(L"MTP streamed reader: GetSize expected 41, got {} hr=0x{:08X}.", sizeBytes, static_cast<unsigned long>(sizeHr)));
    state.Require(readCounter("readFileCalls", readFileCalls, L"after GetSize"), L"MTP streamed reader: failed to read size counter.");
    state.Require(readFileCalls == 0u, std::format(L"MTP streamed reader: GetSize read file contents; readFileCalls={}.", readFileCalls));

    char firstBytes[7] = {};
    unsigned long bytesRead = 0;
    const HRESULT readHr = reader->Read(firstBytes, static_cast<unsigned long>(std::size(firstBytes)), &bytesRead);
    state.Require(SUCCEEDED(readHr) && bytesRead == static_cast<unsigned long>(std::size(firstBytes)) &&
                      std::string_view(firstBytes, bytesRead) == "RedSala",
                  std::format(L"MTP streamed reader: first small read failed. bytes={} hr=0x{:08X}.", bytesRead, static_cast<unsigned long>(readHr)));

    uint64_t lastReadBytes = 0;
    state.Require(readCounter("readFileCalls", readFileCalls, L"after first read"), L"MTP streamed reader: failed to read post-read counter.");
    state.Require(readCounter("lastReadBytes", lastReadBytes, L"after first read"), L"MTP streamed reader: failed to read post-read byte counter.");
    state.Require(readFileCalls == 1u, std::format(L"MTP streamed reader: expected one backend read after first Read, got {}.", readFileCalls));
    state.Require(lastReadBytes == static_cast<uint64_t>(std::size(firstBytes)),
                  std::format(L"MTP streamed reader: first backend read should be bounded to {} bytes, got {}.",
                              static_cast<uint64_t>(std::size(firstBytes)),
                              lastReadBytes));

    return state.failure.empty();
});

SelfTest::RunCase(options,
                  suite,
                  L"mtp_concurrency_is_serialized",
                  [&](SelfTest::CaseState& state) noexcept
{
    CreatedFileSystemInstance created;
    const HRESULT createHr = TryCreateFakeMtpFileSystemInstance(R"json({"operationDelayMs":25})json", R"json({"readOnly":true})json", L"/", created);
    if (createHr == HRESULT_FROM_WIN32(ERROR_PROC_NOT_FOUND))
    {
        return state.Skip(L"MTP fake backend export is available only in debug plugin builds.");
    }
    state.Require(SUCCEEDED(createHr) && created.fileSystem,
                  std::format(L"MTP concurrency: create selftest instance failed. hr=0x{:08X}", static_cast<unsigned long>(createHr)));
    if (FAILED(createHr) || ! created.fileSystem)
    {
        return false;
    }

    wil::com_ptr<IFileSystemIO> io;
    state.Require(CreateFileSystemIo(created.fileSystem, io), L"MTP concurrency: missing IFileSystemIO.");
    if (! io)
    {
        return false;
    }

    constexpr std::wstring_view kPhotoPath       = L"/Fake Phone [devpuid:fake-device]/Internal Storage/DCIM/Camera/photo001.txt";
    constexpr std::string_view kExpectedContents = "RedSalamander deterministic MTP fixture\r\n";
    constexpr size_t kWorkerCount                = 8u;

    const auto readFixtureNoReport = [&](std::string& out) noexcept
    {
        out.clear();
        wil::com_ptr<IFileReader> reader;
        HRESULT hr = io->CreateFileReader(kPhotoPath.data(), reader.put());
        if (FAILED(hr) || ! reader)
        {
            return FAILED(hr) ? hr : E_FAIL;
        }

        uint64_t sizeBytes = 0;
        hr                 = reader->GetSize(&sizeBytes);
        if (FAILED(hr))
        {
            return hr;
        }

        out.resize(static_cast<size_t>(sizeBytes));
        unsigned long bytesRead = 0;
        hr                      = reader->Read(out.empty() ? nullptr : out.data(), static_cast<unsigned long>(out.size()), &bytesRead);
        if (FAILED(hr))
        {
            return hr;
        }
        out.resize(bytesRead);
        return S_OK;
    };

    std::atomic_bool start{false};
    std::atomic_uint32_t successCount{0};
    std::atomic_uint32_t failureCount{0};
    {
        std::array<std::jthread, kWorkerCount> workers;
        for (std::jthread& worker : workers)
        {
            worker = std::jthread([&]() noexcept
            {
                while (! start.load(std::memory_order_acquire))
                {
                    std::this_thread::yield();
                }

                std::string readBack;
                const HRESULT hr = readFixtureNoReport(readBack);
                if (SUCCEEDED(hr) && readBack == kExpectedContents)
                {
                    static_cast<void>(successCount.fetch_add(1u, std::memory_order_acq_rel));
                }
                else
                {
                    static_cast<void>(failureCount.fetch_add(1u, std::memory_order_acq_rel));
                }
            });
        }

        start.store(true, std::memory_order_release);
    }

    state.Require(successCount.load(std::memory_order_acquire) == kWorkerCount && failureCount.load(std::memory_order_acquire) == 0u,
                  std::format(L"MTP concurrency: expected {} successful concurrent reads, got success={} failure={}.",
                              kWorkerCount,
                              successCount.load(std::memory_order_acquire),
                              failureCount.load(std::memory_order_acquire)));

    const char* properties = nullptr;
    const HRESULT propsHr  = io->GetItemProperties(kPhotoPath.data(), &properties);
    state.Require(SUCCEEDED(propsHr) && properties != nullptr,
                  std::format(L"MTP concurrency: GetItemProperties failed. hr=0x{:08X}", static_cast<unsigned long>(propsHr)));
    if (SUCCEEDED(propsHr) && properties)
    {
        const std::string_view props(properties);
        state.Require(props.find(R"json("operationDelayMs":25)json") != std::string_view::npos,
                      L"MTP concurrency: fake backend did not report the configured operation delay.");
        state.Require(extractJsonUInt(props, "maxConcurrentBackendCalls") == 1u,
                      L"MTP concurrency: backend observed concurrent device calls; expected maxConcurrentBackendCalls=1.");
    }

    return state.failure.empty();
});

SelfTest::RunCase(options,
                  suite,
                  L"mtp_backend_command_worker_is_reused",
                  [&](SelfTest::CaseState& state) noexcept
{
    CreatedFileSystemInstance created;
    const HRESULT createHr = TryCreateFakeMtpFileSystemInstance(R"json({"operationDelayMs":25})json", R"json({"readOnly":true})json", L"/", created);
    if (createHr == HRESULT_FROM_WIN32(ERROR_PROC_NOT_FOUND))
    {
        return state.Skip(L"MTP fake backend export is available only in debug plugin builds.");
    }
    state.Require(SUCCEEDED(createHr) && created.fileSystem,
                  std::format(L"MTP backend worker reuse: create selftest instance failed. hr=0x{:08X}", static_cast<unsigned long>(createHr)));
    if (FAILED(createHr) || ! created.fileSystem)
    {
        return false;
    }

    wil::com_ptr<IFileSystemIO> io;
    state.Require(CreateFileSystemIo(created.fileSystem, io), L"MTP backend worker reuse: missing IFileSystemIO.");
    if (! io)
    {
        return false;
    }

    constexpr std::wstring_view kPhotoPath = L"/Fake Phone [devpuid:fake-device]/Internal Storage/DCIM/Camera/photo001.txt";
    constexpr size_t kWorkerCount          = 4u;
    std::atomic_bool start{false};
    std::atomic_uint32_t successCount{0};
    std::atomic_uint32_t failureCount{0};
    {
        std::array<std::jthread, kWorkerCount> workers;
        for (std::jthread& worker : workers)
        {
            worker = std::jthread([&]() noexcept
            {
                while (! start.load(std::memory_order_acquire))
                {
                    std::this_thread::yield();
                }

                unsigned long attributes = 0;
                const HRESULT hr         = io->GetAttributes(kPhotoPath.data(), &attributes);
                if (SUCCEEDED(hr) && (attributes & FILE_ATTRIBUTE_DIRECTORY) == 0)
                {
                    static_cast<void>(successCount.fetch_add(1u, std::memory_order_acq_rel));
                }
                else
                {
                    static_cast<void>(failureCount.fetch_add(1u, std::memory_order_acq_rel));
                }
            });
        }

        start.store(true, std::memory_order_release);
    }

    state.Require(successCount.load(std::memory_order_acquire) == kWorkerCount && failureCount.load(std::memory_order_acquire) == 0u,
                  std::format(L"MTP backend worker reuse: expected {} successful concurrent attributes calls, got success={} failure={}.",
                              kWorkerCount,
                              successCount.load(std::memory_order_acquire),
                              failureCount.load(std::memory_order_acquire)));

    const char* properties = nullptr;
    const HRESULT propsHr  = io->GetItemProperties(kPhotoPath.data(), &properties);
    state.Require(SUCCEEDED(propsHr) && properties != nullptr,
                  std::format(L"MTP backend worker reuse: GetItemProperties failed. hr=0x{:08X}", static_cast<unsigned long>(propsHr)));
    if (FAILED(propsHr) || properties == nullptr)
    {
        return false;
    }

    const std::string_view props(properties);
    const std::optional<uint64_t> backendThreadIdsObserved = extractJsonUInt(props, "backendThreadIdsObserved");
    state.Require(backendThreadIdsObserved.has_value(), L"MTP backend worker reuse: fake instrumentation omitted backendThreadIdsObserved.");
    if (backendThreadIdsObserved.has_value())
    {
        state.Require(backendThreadIdsObserved.value() == 1u,
                      std::format(L"MTP backend worker reuse: expected one long-lived backend worker thread, observed {}.",
                                  backendThreadIdsObserved.value()));
    }
    state.Require(props.find(R"json("backendThreadIdsOverflow":false)json") != std::string_view::npos,
                  L"MTP backend worker reuse: backend thread-id instrumentation overflowed.");
    state.Require(extractJsonUInt(props, "maxConcurrentBackendCalls") == 1u,
                  L"MTP backend worker reuse: backend calls were not serialized while reusing the worker.");

    return state.failure.empty();
});

SelfTest::RunCase(options,
                  suite,
                  L"mtp_wpd_session_and_path_cache_reuse",
                  [&](SelfTest::CaseState& state) noexcept
{
    CreatedFileSystemInstance created;
    const HRESULT createHr = TryCreateWpdCacheMtpFileSystemInstance("{}", R"json({"readOnly":true})json", L"/", created);
    if (createHr == HRESULT_FROM_WIN32(ERROR_PROC_NOT_FOUND))
    {
        return state.Skip(L"MTP WPD-cache fixture export is available only in debug plugin builds.");
    }
    state.Require(SUCCEEDED(createHr) && created.fileSystem,
                  std::format(L"MTP WPD cache: create selftest instance failed. hr=0x{:08X}", static_cast<unsigned long>(createHr)));
    if (FAILED(createHr) || ! created.fileSystem)
    {
        return false;
    }

    wil::com_ptr<IFileSystemIO> io;
    state.Require(CreateFileSystemIo(created.fileSystem, io), L"MTP WPD cache: missing IFileSystemIO.");
    if (! io)
    {
        return false;
    }

    constexpr std::wstring_view kPhotoPath = L"/Fake Phone [devid:000000000000F00D]/Internal Storage/DCIM/Camera/photo001.txt";
    for (uint32_t callIndex = 0; callIndex < 4u; ++callIndex)
    {
        unsigned long attributes = 0;
        const HRESULT attrsHr    = io->GetAttributes(kPhotoPath.data(), &attributes);
        state.Require(SUCCEEDED(attrsHr) && (attributes & FILE_ATTRIBUTE_DIRECTORY) == 0,
                      std::format(L"MTP WPD cache: GetAttributes call {} failed. hr=0x{:08X}", callIndex, static_cast<unsigned long>(attrsHr)));
    }

    const char* properties = nullptr;
    const HRESULT propsHr  = io->GetItemProperties(kPhotoPath.data(), &properties);
    state.Require(SUCCEEDED(propsHr) && properties != nullptr,
                  std::format(L"MTP WPD cache: GetItemProperties failed. hr=0x{:08X}", static_cast<unsigned long>(propsHr)));
    if (FAILED(propsHr) || properties == nullptr)
    {
        return false;
    }

    const std::string_view props(properties);
    state.Require(props.find(R"json("backend":"wpd-selftest")json") != std::string_view::npos,
                  L"MTP WPD cache: properties did not identify the WPD-shaped selftest backend.");

    const std::optional<uint64_t> deviceEnumerations = extractJsonUInt(props, "deviceEnumerations");
    const std::optional<uint64_t> sessionOpens       = extractJsonUInt(props, "sessionOpens");
    const std::optional<uint64_t> childEnumerations  = extractJsonUInt(props, "childResolveEnumerations");
    const std::optional<uint64_t> pathCacheHits      = extractJsonUInt(props, "pathCacheHits");
    state.Require(deviceEnumerations.has_value(), L"MTP WPD cache: missing deviceEnumerations instrumentation.");
    state.Require(sessionOpens.has_value(), L"MTP WPD cache: missing sessionOpens instrumentation.");
    state.Require(childEnumerations.has_value(), L"MTP WPD cache: missing childResolveEnumerations instrumentation.");
    state.Require(pathCacheHits.has_value(), L"MTP WPD cache: missing pathCacheHits instrumentation.");

    if (deviceEnumerations.has_value())
    {
        state.Require(deviceEnumerations.value() == 1u,
                      std::format(L"MTP WPD cache: expected one device enumeration, got {}.", deviceEnumerations.value()));
    }
    if (sessionOpens.has_value())
    {
        state.Require(sessionOpens.value() == 1u, std::format(L"MTP WPD cache: expected one WPD session open, got {}.", sessionOpens.value()));
    }
    if (childEnumerations.has_value())
    {
        state.Require(childEnumerations.value() == 4u,
                      std::format(L"MTP WPD cache: expected one four-segment ancestor walk, got {} child enumerations.",
                                  childEnumerations.value()));
    }
    if (pathCacheHits.has_value())
    {
        state.Require(pathCacheHits.value() >= 4u,
                      std::format(L"MTP WPD cache: expected repeated path cache hits, got {}.", pathCacheHits.value()));
    }

    return state.failure.empty();
});

SelfTest::RunCase(options,
                  suite,
                  L"mtp_wpd_cache_failure_reopens_session_and_refreshes_size",
                  [&](SelfTest::CaseState& state) noexcept
{
    constexpr std::wstring_view kPhotoPath = L"/Fake Phone [devid:000000000000F00D]/Internal Storage/DCIM/Camera/photo001.txt";

    CreatedFileSystemInstance unsupported;
    const HRESULT unsupportedHr = TryCreateWpdCacheMtpFileSystemInstance(
        R"json({"unsupportedFixtureOption":true})json", R"json({"readOnly":true})json", L"/", unsupported);
    if (unsupportedHr == HRESULT_FROM_WIN32(ERROR_PROC_NOT_FOUND))
    {
        return state.Skip(L"MTP WPD-cache fixture export is available only in debug plugin builds.");
    }
    state.Require(unsupportedHr == E_INVALIDARG,
                  std::format(L"MTP WPD fixture options: unsupported option expected E_INVALIDARG, got 0x{:08X}.",
                              static_cast<unsigned long>(unsupportedHr)));

    CreatedFileSystemInstance sessionFailure;
    HRESULT hr = TryCreateWpdCacheMtpFileSystemInstance(
        R"json({"sessionDeathOnce":true})json", R"json({"readOnly":true})json", L"/", sessionFailure);
    state.Require(SUCCEEDED(hr) && sessionFailure.fileSystem,
                  std::format(L"MTP WPD session recovery: fixture creation failed. hr=0x{:08X}.", static_cast<unsigned long>(hr)));
    if (FAILED(hr) || ! sessionFailure.fileSystem)
    {
        return false;
    }

    wil::com_ptr<IFileSystemIO> sessionIo;
    state.Require(CreateFileSystemIo(sessionFailure.fileSystem, sessionIo), L"MTP WPD session recovery: missing IFileSystemIO.");
    if (! sessionIo)
    {
        return false;
    }
    unsigned long attributes = 0;
    const HRESULT firstHr    = sessionIo->GetAttributes(kPhotoPath.data(), &attributes);
    state.Require(firstHr == RPC_E_DISCONNECTED,
                  std::format(L"MTP WPD session recovery: first injected session death expected RPC_E_DISCONNECTED, got 0x{:08X}.",
                              static_cast<unsigned long>(firstHr)));
    const HRESULT retryHr = sessionIo->GetAttributes(kPhotoPath.data(), &attributes);
    state.Require(SUCCEEDED(retryHr) && (attributes & FILE_ATTRIBUTE_DIRECTORY) == 0,
                  std::format(L"MTP WPD session recovery: retry did not reopen a fresh session. hr=0x{:08X}.",
                              static_cast<unsigned long>(retryHr)));
    const char* sessionProperties = nullptr;
    hr = sessionIo->GetItemProperties(kPhotoPath.data(), &sessionProperties);
    state.Require(SUCCEEDED(hr) && sessionProperties != nullptr && std::string_view(sessionProperties).find(R"json("sessionOpens":2)json") != std::string_view::npos,
                  L"MTP WPD session recovery: instrumentation did not record the fresh session open.");

    CreatedFileSystemInstance changingSize;
    hr = TryCreateWpdCacheMtpFileSystemInstance(
        R"json({"changeFileSizeAfterFirstLookup":true,"readFileDelayMs":1})json", R"json({"readOnly":true})json", L"/", changingSize);
    state.Require(SUCCEEDED(hr) && changingSize.fileSystem,
                  std::format(L"MTP WPD metadata refresh: fixture creation failed. hr=0x{:08X}.", static_cast<unsigned long>(hr)));
    if (FAILED(hr) || ! changingSize.fileSystem)
    {
        return false;
    }

    wil::com_ptr<IFileSystemIO> changingIo;
    state.Require(CreateFileSystemIo(changingSize.fileSystem, changingIo), L"MTP WPD metadata refresh: missing IFileSystemIO.");
    if (! changingIo)
    {
        return false;
    }
    attributes = 0;
    hr         = changingIo->GetAttributes(kPhotoPath.data(), &attributes);
    state.Require(SUCCEEDED(hr), L"MTP WPD metadata refresh: initial metadata lookup failed.");

    wil::com_ptr<IFileReader> reader;
    hr = changingIo->CreateFileReader(kPhotoPath.data(), reader.put());
    state.Require(SUCCEEDED(hr) && reader,
                  std::format(L"MTP WPD metadata refresh: CreateFileReader failed. hr=0x{:08X}.", static_cast<unsigned long>(hr)));
    if (! reader)
    {
        return false;
    }
    constexpr std::string_view kChangedPayload = "RedSalamander refreshed MTP fixture payload\r\n";
    unsigned long long sizeBytes = 0;
    hr = reader->GetSize(&sizeBytes);
    state.Require(SUCCEEDED(hr) && sizeBytes == kChangedPayload.size(),
                  std::format(L"MTP WPD metadata refresh: reader retained stale size {}; expected {}. hr=0x{:08X}.",
                              sizeBytes,
                              kChangedPayload.size(),
                              static_cast<unsigned long>(hr)));
    std::vector<char> payload(kChangedPayload.size());
    unsigned long bytesRead = 0;
    hr = reader->Read(payload.data(), static_cast<unsigned long>(payload.size()), &bytesRead);
    state.Require(SUCCEEDED(hr) && bytesRead == payload.size() && std::string_view(payload.data(), bytesRead) == kChangedPayload,
                  std::format(L"MTP WPD metadata refresh: refreshed payload read failed. bytes={} hr=0x{:08X}.",
                              bytesRead,
                              static_cast<unsigned long>(hr)));

    return state.failure.empty();
});

SelfTest::RunCase(options,
                  suite,
                  L"mtp_overwrite_journal_generation_and_absent_cache_are_constant_cost",
                  [&](SelfTest::CaseState& state) noexcept
{
    CreatedFileSystemInstance created;
    const HRESULT createHr = TryCreateFakeMtpFileSystemInstance("{}", R"json({"readOnly":false,"host":"firebreak-journal-cache"})json", L"/", created);
    if (createHr == HRESULT_FROM_WIN32(ERROR_PROC_NOT_FOUND))
    {
        return state.Skip(L"MTP fake backend and journal test exports are available only in debug plugin builds.");
    }
    state.Require(SUCCEEDED(createHr) && created.fileSystem && created.module,
                  std::format(L"MTP journal cache: fixture creation failed. hr=0x{:08X}.", static_cast<unsigned long>(createHr)));
    if (FAILED(createHr) || ! created.fileSystem || ! created.module)
    {
        return false;
    }

    using RunGenerationSelfTestFunc = BOOL(__stdcall*)() noexcept;
    using JournalProbeCountFunc     = uint64_t(__stdcall*)() noexcept;
#pragma warning(push)
#pragma warning(disable : 4191) // unsafe conversion from FARPROC
    const auto runGenerationSelfTest =
        reinterpret_cast<RunGenerationSelfTestFunc>(GetProcAddress(created.module.get(), "RedSalamanderMtpRunJournalGenerationSelfTest"));
    const auto resetProbeCount =
        reinterpret_cast<JournalProbeCountFunc>(GetProcAddress(created.module.get(), "RedSalamanderMtpResetJournalProbeCountForSelfTest"));
    const auto getProbeCount =
        reinterpret_cast<JournalProbeCountFunc>(GetProcAddress(created.module.get(), "RedSalamanderMtpGetJournalProbeCountForSelfTest"));
#pragma warning(pop)
    state.Require(runGenerationSelfTest && resetProbeCount && getProbeCount, L"MTP journal cache: required debug exports are missing.");
    if (! runGenerationSelfTest || ! resetProbeCount || ! getProbeCount)
    {
        return false;
    }
    state.Require(runGenerationSelfTest() == TRUE,
                  L"MTP journal cache: a stale absent observation was accepted after another instance began a journal write.");

    uint64_t baselineProbes = 0u;
    for (uint32_t index = 0u; index < 16u; ++index)
    {
        static_cast<void>(resetProbeCount());
        const std::wstring missingPath = std::format(
            L"/Fake Phone [devpuid:fake-device]/Internal Storage/DCIM/Camera/firebreak-baseline-missing-{}.txt", index);
        const HRESULT deleteHr = created.fileSystem->DeleteItem(missingPath.c_str(), FILESYSTEM_FLAG_NONE, nullptr, nullptr, nullptr);
        state.Require(deleteHr == HRESULT_FROM_WIN32(ERROR_FILE_NOT_FOUND),
                      std::format(L"MTP journal cache baseline: missing delete {} returned 0x{:08X}.", index, static_cast<unsigned long>(deleteHr)));
        baselineProbes += getProbeCount();
    }
    Debug::Perf::EmitValue(L"mtp.overwrite.journal_filesystem_probes.baseline_no_cache", baselineProbes, S_OK);

    static_cast<void>(resetProbeCount());
    for (uint32_t index = 0u; index < 16u; ++index)
    {
        const std::wstring missingPath = std::format(
            L"/Fake Phone [devpuid:fake-device]/Internal Storage/DCIM/Camera/firebreak-missing-{}.txt", index);
        const HRESULT deleteHr = created.fileSystem->DeleteItem(missingPath.c_str(), FILESYSTEM_FLAG_NONE, nullptr, nullptr, nullptr);
        state.Require(deleteHr == HRESULT_FROM_WIN32(ERROR_FILE_NOT_FOUND),
                      std::format(L"MTP journal cache: missing delete {} returned 0x{:08X}.", index, static_cast<unsigned long>(deleteHr)));
    }
    const uint64_t probes = getProbeCount();
    state.Require(probes <= 1u,
                  std::format(L"MTP journal cache: 16 deletes performed {} filesystem journal probes; expected O(1) (at most one).", probes));
    Debug::Perf::EmitValue(L"mtp.overwrite.journal_filesystem_probes.candidate", probes, S_OK);
    state.Require(baselineProbes == 16u && probes < baselineProbes,
                  std::format(L"MTP journal cache: deterministic baseline/candidate probes were {}/{}; expected 16/at-most-1.",
                              baselineProbes,
                              probes));

    return state.failure.empty();
});

SelfTest::RunCase(options,
                  suite,
                  L"mtp_property_fetch_is_batched",
                  [&](SelfTest::CaseState& state) noexcept
{
    CreatedFileSystemInstance created;
    const HRESULT createHr = TryCreateFakeMtpFileSystemInstance("{}", R"json({"readOnly":true})json", L"/", created);
    if (createHr == HRESULT_FROM_WIN32(ERROR_PROC_NOT_FOUND))
    {
        return state.Skip(L"MTP fake backend export is available only in debug plugin builds.");
    }
    state.Require(SUCCEEDED(createHr) && created.fileSystem,
                  std::format(L"MTP property batching: create selftest instance failed. hr=0x{:08X}", static_cast<unsigned long>(createHr)));
    if (FAILED(createHr) || ! created.fileSystem)
    {
        return false;
    }

    wil::com_ptr<IFileSystemIO> io;
    state.Require(CreateFileSystemIo(created.fileSystem, io), L"MTP property batching: missing IFileSystemIO.");
    if (! io)
    {
        return false;
    }

    constexpr std::wstring_view kCameraPath = L"/Fake Phone [devpuid:fake-device]/Internal Storage/DCIM/Camera";
    const auto cameraEntries                = SnapshotDirectoryEntries(created.fileSystem, kCameraPath.data(), state, L"MTP property batching camera");
    state.Require(FindDirectoryEntrySnapshot(cameraEntries, L"photo001.txt") != nullptr,
                  L"MTP property batching: expected photo001.txt after camera enumeration.");
    state.Require(FindDirectoryEntrySnapshot(cameraEntries, L"name [puid:literal].txt") != nullptr,
                  L"MTP property batching: expected literal suffix fixture after camera enumeration.");

    const char* properties                 = nullptr;
    constexpr std::wstring_view kPhotoPath = L"/Fake Phone [devpuid:fake-device]/Internal Storage/DCIM/Camera/photo001.txt";
    const HRESULT propsHr                  = io->GetItemProperties(kPhotoPath.data(), &properties);
    state.Require(SUCCEEDED(propsHr) && properties != nullptr,
                  std::format(L"MTP property batching: GetItemProperties failed. hr=0x{:08X}", static_cast<unsigned long>(propsHr)));
    if (SUCCEEDED(propsHr) && properties)
    {
        const std::string_view props(properties);
        state.Require(props.find(R"json("propertyBatchCalls":1)json") != std::string_view::npos,
                      L"MTP property batching: directory enumeration did not use exactly one logical property batch.");
        state.Require(props.find(R"json("propertyPerItemCalls":0)json") != std::string_view::npos,
                      L"MTP property batching: directory enumeration used per-item property calls.");
    }

    return state.failure.empty();
});

SelfTest::RunCase(options,
                  suite,
                  L"mtp_public_writer_stages_until_commit",
                  [&](SelfTest::CaseState& state) noexcept
{
    CreatedFileSystemInstance created;
    const HRESULT createHr = TryCreateFakeMtpFileSystemInstance("{}", R"json({"readOnly":false})json", L"/", created);
    if (createHr == HRESULT_FROM_WIN32(ERROR_PROC_NOT_FOUND))
    {
        return state.Skip(L"MTP fake backend export is available only in debug plugin builds.");
    }
    state.Require(SUCCEEDED(createHr) && created.fileSystem,
                  std::format(L"MTP staged writer: create selftest instance failed. hr=0x{:08X}", static_cast<unsigned long>(createHr)));
    if (FAILED(createHr) || ! created.fileSystem)
    {
        return false;
    }

    wil::com_ptr<IFileSystemIO> io;
    state.Require(CreateFileSystemIo(created.fileSystem, io), L"MTP staged writer: missing IFileSystemIO.");
    if (! io)
    {
        return false;
    }

    constexpr std::wstring_view kProbePath = L"/Fake Phone [devpuid:fake-device]/Internal Storage/DCIM/Camera/photo001.txt";
    const auto widenAscii                  = [](std::string_view text) { return std::wstring(text.begin(), text.end()); };
    const auto requireWriteCallCount       = [&](uint32_t expected, std::wstring_view label) noexcept
    {
        const char* properties = nullptr;
        const HRESULT propsHr  = io->GetItemProperties(kProbePath.data(), &properties);
        state.Require(SUCCEEDED(propsHr) && properties != nullptr,
                      std::format(L"MTP staged writer: {} GetItemProperties failed. hr=0x{:08X}", label, static_cast<unsigned long>(propsHr)));
        if (FAILED(propsHr) || properties == nullptr)
        {
            return;
        }

        const std::string expectedToken = std::format(R"json("writeFileCalls":{})json", expected);
        state.Require(std::string_view(properties).find(expectedToken) != std::string_view::npos,
                      std::format(L"MTP staged writer: {} expected token {} in fake backend properties.", label, widenAscii(expectedToken)));
    };

    const std::wstring guid = MakeGuidText();
    state.Require(! guid.empty(), L"MTP staged writer: failed to generate unique file name.");
    if (guid.empty())
    {
        return false;
    }

    const std::wstring baseFolder    = L"/Fake Phone [devpuid:fake-device]/Internal Storage/DCIM/Camera";
    const std::wstring abortedPath   = baseFolder + L"/aborted-writer-" + guid + L".txt";
    const std::wstring committedPath = baseFolder + L"/committed-writer-" + guid + L".txt";
    auto cleanup                     = wil::scope_exit([&]() noexcept
    {
        static_cast<void>(created.fileSystem->DeleteItem(abortedPath.c_str(), FILESYSTEM_FLAG_NONE, nullptr, nullptr, nullptr));
        static_cast<void>(created.fileSystem->DeleteItem(committedPath.c_str(), FILESYSTEM_FLAG_NONE, nullptr, nullptr, nullptr));
    });

    requireWriteCallCount(0u, L"initial");

    {
        wil::com_ptr<IFileWriter> writer;
        const HRESULT writerHr = io->CreateFileWriter(abortedPath.c_str(), FILESYSTEM_FLAG_NONE, writer.put());
        state.Require(SUCCEEDED(writerHr) && writer,
                      std::format(L"MTP staged writer: CreateFileWriter abort path failed. hr=0x{:08X}", static_cast<unsigned long>(writerHr)));
        if (FAILED(writerHr) || ! writer)
        {
            return false;
        }

        constexpr std::string_view kAbortPayload = "abort before commit";
        const unsigned long abortPayloadBytes    = static_cast<unsigned long>(kAbortPayload.size());
        unsigned long written                    = 0;
        const HRESULT writeHr                    = writer->Write(kAbortPayload.data(), abortPayloadBytes, &written);
        state.Require(SUCCEEDED(writeHr) && written == abortPayloadBytes,
                      std::format(L"MTP staged writer: abort-path Write failed. bytes={} hr=0x{:08X}.", written, static_cast<unsigned long>(writeHr)));
        requireWriteCallCount(0u, L"after Write before abort");
    }

    unsigned long attrs        = 0;
    const HRESULT abortAttrsHr = io->GetAttributes(abortedPath.c_str(), &attrs);
    state.Require(abortAttrsHr == HRESULT_FROM_WIN32(ERROR_FILE_NOT_FOUND),
                  std::format(L"MTP staged writer: release without Commit created a backend item. hr=0x{:08X}.", static_cast<unsigned long>(abortAttrsHr)));
    requireWriteCallCount(0u, L"after release without Commit");

    constexpr std::string_view kCommitPayload = "commit staged payload";
    const unsigned long commitPayloadBytes    = static_cast<unsigned long>(kCommitPayload.size());
    wil::com_ptr<IFileWriter> writer;
    const HRESULT writerHr = io->CreateFileWriter(committedPath.c_str(), FILESYSTEM_FLAG_NONE, writer.put());
    state.Require(SUCCEEDED(writerHr) && writer,
                  std::format(L"MTP staged writer: CreateFileWriter commit path failed. hr=0x{:08X}", static_cast<unsigned long>(writerHr)));
    if (FAILED(writerHr) || ! writer)
    {
        return false;
    }

    unsigned long written = 0;
    const HRESULT writeHr = writer->Write(kCommitPayload.data(), commitPayloadBytes, &written);
    state.Require(SUCCEEDED(writeHr) && written == commitPayloadBytes,
                  std::format(L"MTP staged writer: commit-path Write failed. bytes={} hr=0x{:08X}.", written, static_cast<unsigned long>(writeHr)));
    requireWriteCallCount(0u, L"after Write before Commit");

    const HRESULT commitHr = writer->Commit();
    state.Require(SUCCEEDED(commitHr), std::format(L"MTP staged writer: Commit failed. hr=0x{:08X}", static_cast<unsigned long>(commitHr)));
    writer.reset();
    requireWriteCallCount(1u, L"after Commit");

    std::string readBack;
    state.Require(ReadPluginFileText(io.get(), committedPath.c_str(), readBack, state, L"MTP staged writer committed read"),
                  L"MTP staged writer: failed to read committed file.");
    state.Require(readBack == kCommitPayload, L"MTP staged writer: committed file contents did not match.");

    return state.failure.empty();
});

SelfTest::RunCase(options,
                  suite,
                  L"mtp_fake_backend_enumerate_read_and_capabilities",
                  [&](SelfTest::CaseState& state) noexcept
{
    CreatedFileSystemInstance created;
    const HRESULT createHr = TryCreateFakeMtpFileSystemInstance("{}", R"json({"readOnly":false,"byteVerifyOnOverwrite":"deviceReread"})json", L"/", created);
    if (createHr == HRESULT_FROM_WIN32(ERROR_PROC_NOT_FOUND))
    {
        return state.Skip(L"MTP fake backend export is available only in debug plugin builds.");
    }
    state.Require(SUCCEEDED(createHr) && created.fileSystem,
                  std::format(L"MTP fake enumerate: create selftest instance failed. hr=0x{:08X}", static_cast<unsigned long>(createHr)));
    if (FAILED(createHr) || ! created.fileSystem)
    {
        return false;
    }

    wil::com_ptr<IFileSystemIO> io;
    state.Require(CreateFileSystemIo(created.fileSystem, io), L"MTP fake enumerate: missing IFileSystemIO.");

    wil::com_ptr<IFileSystemDirectoryOperations> dirOps;
    state.Require(CreateFileSystemDirectoryOperations(created.fileSystem, dirOps), L"MTP fake enumerate: missing IFileSystemDirectoryOperations.");
    if (! io || ! dirOps)
    {
        return false;
    }

    const auto rootEntries                    = SnapshotDirectoryEntries(created.fileSystem, L"/", state, L"MTP fake enumerate root");
    const DirectoryEntrySnapshot* deviceEntry = FindDirectoryEntrySnapshot(rootEntries, L"Fake Phone [devpuid:fake-device]");
    state.Require(deviceEntry != nullptr, L"MTP fake enumerate: root did not contain the fake device.");
    state.Require(deviceEntry == nullptr || (deviceEntry->attributes & FILE_ATTRIBUTE_DIRECTORY) != 0,
                  L"MTP fake enumerate: fake device was not reported as a directory.");

    constexpr std::wstring_view kCameraPath  = L"/Fake Phone [devpuid:fake-device]/Internal Storage/DCIM/Camera";
    const auto cameraEntries                 = SnapshotDirectoryEntries(created.fileSystem, kCameraPath.data(), state, L"MTP fake enumerate camera");
    const DirectoryEntrySnapshot* photoEntry = FindDirectoryEntrySnapshot(cameraEntries, L"photo001.txt");
    state.Require(photoEntry != nullptr, L"MTP fake enumerate: camera folder did not contain photo001.txt.");
    state.Require(photoEntry == nullptr || (photoEntry->attributes & FILE_ATTRIBUTE_DIRECTORY) == 0,
                  L"MTP fake enumerate: photo001.txt was reported as a directory.");
    state.Require(photoEntry == nullptr || photoEntry->sizeBytes == 41u,
                  std::format(L"MTP fake enumerate: photo001.txt expected 41 bytes, got {}.", photoEntry ? photoEntry->sizeBytes : 0u));
    state.Require(FindDirectoryEntrySnapshot(cameraEntries, L"name [puid:literal].txt") != nullptr,
                  L"MTP fake enumerate: literal [puid:...] filename was not preserved.");

    constexpr std::wstring_view kPhotoPath    = L"/Fake Phone [devpuid:fake-device]/Internal Storage/DCIM/Camera/photo001.txt";
    const auto readFakeInstrumentationCounter = [&](std::string_view key, uint64_t& value, std::wstring_view label) noexcept -> bool
    {
        value                      = 0;
        const char* properties     = nullptr;
        const HRESULT propertiesHr = io->GetItemProperties(kPhotoPath.data(), &properties);
        state.Require(SUCCEEDED(propertiesHr) && properties != nullptr,
                      std::format(L"MTP fake enumerate: {} GetItemProperties failed. hr=0x{:08X}", label, static_cast<unsigned long>(propertiesHr)));
        if (FAILED(propertiesHr) || properties == nullptr)
        {
            return false;
        }

        const std::string token = std::format(R"json("{}":)json", key);
        const std::string_view json(properties);
        size_t position = json.find(token);
        if (position == std::string_view::npos)
        {
            state.Require(false, std::format(L"MTP fake enumerate: {} missing instrumentation counter {}.", label, std::wstring(key.begin(), key.end())));
            return false;
        }

        position += token.size();
        if (position >= json.size() || json[position] < '0' || json[position] > '9')
        {
            state.Require(false, std::format(L"MTP fake enumerate: {} counter {} was not numeric.", label, std::wstring(key.begin(), key.end())));
            return false;
        }

        while (position < json.size() && json[position] >= '0' && json[position] <= '9')
        {
            const uint64_t digit = static_cast<uint64_t>(json[position] - '0');
            if (value > ((std::numeric_limits<uint64_t>::max)() - digit) / 10u)
            {
                state.Require(false, std::format(L"MTP fake enumerate: {} counter {} overflowed.", label, std::wstring(key.begin(), key.end())));
                return false;
            }
            value = (value * 10u) + digit;
            ++position;
        }

        return true;
    };

    std::string readBack;
    state.Require(ReadPluginFileText(io.get(), kPhotoPath.data(), readBack, state, L"MTP fake enumerate read"),
                  L"MTP fake enumerate: failed to read photo fixture.");
    state.Require(readBack == "RedSalamander deterministic MTP fixture\r\n", L"MTP fake enumerate: photo fixture contents did not match.");

    wil::com_ptr<IFileReader> reader;
    const HRESULT readerHr = io->CreateFileReader(kPhotoPath.data(), reader.put());
    state.Require(SUCCEEDED(readerHr) && reader,
                  std::format(L"MTP fake enumerate: CreateFileReader for seek failed. hr=0x{:08X}", static_cast<unsigned long>(readerHr)));
    if (SUCCEEDED(readerHr) && reader)
    {
        uint64_t newPosition = 0;
        const HRESULT seekHr = reader->Seek(14, FILE_BEGIN, &newPosition);
        state.Require(SUCCEEDED(seekHr) && newPosition == 14u,
                      std::format(L"MTP fake enumerate: seek expected position 14, got {} hr=0x{:08X}.", newPosition, static_cast<unsigned long>(seekHr)));

        char chunk[13]                         = {};
        const unsigned long expectedChunkBytes = static_cast<unsigned long>(std::size(chunk));
        unsigned long chunkRead                = 0;
        const HRESULT readHr                   = reader->Read(chunk, expectedChunkBytes, &chunkRead);
        state.Require(
            SUCCEEDED(readHr) && chunkRead == expectedChunkBytes,
            std::format(
                L"MTP fake enumerate: seek read expected {} bytes, got {} hr=0x{:08X}.", expectedChunkBytes, chunkRead, static_cast<unsigned long>(readHr)));
        state.Require(std::string_view(chunk, chunkRead) == "deterministic", L"MTP fake enumerate: seek read returned the wrong slice.");
    }

    FileSystemBasicInformation basicInfo{};
    basicInfo.sizeBytes       = sizeof(FileSystemBasicInformation);
    const HRESULT basicInfoHr = io->GetFileBasicInformation(kPhotoPath.data(), &basicInfo);
    state.Require(SUCCEEDED(basicInfoHr),
                  std::format(L"MTP fake enumerate: GetFileBasicInformation failed. hr=0x{:08X}", static_cast<unsigned long>(basicInfoHr)));
    state.Require((basicInfo.attributes & FILE_ATTRIBUTE_DIRECTORY) == 0, L"MTP fake enumerate: file basic info reported a directory.");
    state.Require(basicInfo.lastWriteTime != 0, L"MTP fake enumerate: file basic info did not include lastWriteTime.");

    const char* itemProperties = nullptr;
    const HRESULT propsHr      = io->GetItemProperties(kPhotoPath.data(), &itemProperties);
    state.Require(SUCCEEDED(propsHr) && itemProperties != nullptr,
                  std::format(L"MTP fake enumerate: GetItemProperties failed. hr=0x{:08X}", static_cast<unsigned long>(propsHr)));
    if (SUCCEEDED(propsHr) && itemProperties)
    {
        const std::string_view props(itemProperties);
        state.Require(props.find(R"json("backend":"fake")json") != std::string_view::npos, L"MTP fake enumerate: properties did not identify fake backend.");
        state.Require(props.find(R"json("persistentId":"file-photo001")json") != std::string_view::npos,
                      L"MTP fake enumerate: properties did not include the deterministic persistent id.");
    }

    uint64_t readFileCallsBeforeDirectorySize = 0;
    uint64_t fileSizeCallsBeforeDirectorySize = 0;
    state.Require(readFakeInstrumentationCounter("readFileCalls", readFileCallsBeforeDirectorySize, L"before directory size"),
                  L"MTP fake enumerate: failed to read baseline readFileCalls.");
    state.Require(readFakeInstrumentationCounter("fileSizeCalls", fileSizeCallsBeforeDirectorySize, L"before directory size"),
                  L"MTP fake enumerate: failed to read baseline fileSizeCalls.");

    FileSystemDirectorySizeResult singleFileSizeResult{};
    singleFileSizeResult.sizeBytes = sizeof(FileSystemDirectorySizeResult);
    const HRESULT singleFileSizeHr = dirOps->GetDirectorySize(kPhotoPath.data(), FILESYSTEM_FLAG_NONE, nullptr, nullptr, &singleFileSizeResult);
    state.Require(SUCCEEDED(singleFileSizeHr) && SUCCEEDED(singleFileSizeResult.status),
                  std::format(L"MTP fake enumerate: single-file GetDirectorySize failed. hr=0x{:08X} status=0x{:08X}",
                              static_cast<unsigned long>(singleFileSizeHr),
                              static_cast<unsigned long>(singleFileSizeResult.status)));
    state.Require(singleFileSizeResult.fileCount == 1u,
                  std::format(L"MTP fake enumerate: single-file GetDirectorySize expected 1 file, got {}.", singleFileSizeResult.fileCount));
    state.Require(singleFileSizeResult.totalBytes == 41u,
                  std::format(L"MTP fake enumerate: single-file GetDirectorySize expected 41 bytes, got {}.", singleFileSizeResult.totalBytes));

    const char* capabilities = nullptr;
    const HRESULT capsHr     = created.fileSystem->GetCapabilities(&capabilities);
    state.Require(SUCCEEDED(capsHr) && capabilities != nullptr,
                  std::format(L"MTP fake enumerate: GetCapabilities failed. hr=0x{:08X}", static_cast<unsigned long>(capsHr)));
    if (SUCCEEDED(capsHr) && capabilities)
    {
        const std::string_view caps(capabilities);
        state.Require(caps.find(R"json("write": true)json") != std::string_view::npos, L"MTP fake enumerate: capabilities did not advertise write.");
        state.Require(caps.find(R"json("pathIdentity")json") != std::string_view::npos, L"MTP fake enumerate: capabilities omitted pathIdentity.");
        state.Require(caps.find(R"json("byteVerifyOnOverwrite": "deviceReread")json") != std::string_view::npos,
                      L"MTP fake enumerate: capabilities did not round-trip byteVerifyOnOverwrite.");
        state.Require(caps.find(R"json("backend": "fake")json") != std::string_view::npos, L"MTP fake enumerate: capabilities did not identify fake backend.");
    }

    FileSystemTransferHints hints{};
    hints.sizeBytes       = sizeof(FileSystemTransferHints);
    const HRESULT hintsHr = created.fileSystem->GetTransferHints(kPhotoPath.data(), FILESYSTEM_COPY, FILESYSTEM_TRANSFER_SOURCE_READ, &hints);
    state.Require(SUCCEEDED(hintsHr), std::format(L"MTP fake enumerate: GetTransferHints failed. hr=0x{:08X}", static_cast<unsigned long>(hintsHr)));
    state.Require(hints.latencyClass == FILESYSTEM_TRANSFER_LATENCY_CLOUD,
                  L"MTP fake enumerate: transfer hints did not classify MTP as high-latency virtual I/O.");
    state.Require((hints.flags & FILESYSTEM_TRANSFER_HINT_HIGH_METADATA_COST) != 0, L"MTP fake enumerate: transfer hints missed high metadata cost.");

    FileSystemDirectorySizeResult sizeResult{};
    sizeResult.sizeBytes      = sizeof(FileSystemDirectorySizeResult);
    const HRESULT directoryHr = dirOps->GetDirectorySize(kCameraPath.data(), FILESYSTEM_FLAG_RECURSIVE, nullptr, nullptr, &sizeResult);
    state.Require(SUCCEEDED(directoryHr) && SUCCEEDED(sizeResult.status),
                  std::format(L"MTP fake enumerate: GetDirectorySize failed. hr=0x{:08X} status=0x{:08X}",
                              static_cast<unsigned long>(directoryHr),
                              static_cast<unsigned long>(sizeResult.status)));
    state.Require(sizeResult.fileCount == 8u, std::format(L"MTP fake enumerate: expected 8 files, got {}.", sizeResult.fileCount));
    state.Require(sizeResult.totalBytes == 194u, std::format(L"MTP fake enumerate: expected 194 total bytes, got {}.", sizeResult.totalBytes));
    uint64_t readFileCallsAfterDirectorySize = 0;
    uint64_t fileSizeCallsAfterDirectorySize = 0;
    state.Require(readFakeInstrumentationCounter("readFileCalls", readFileCallsAfterDirectorySize, L"after directory size"),
                  L"MTP fake enumerate: failed to read final readFileCalls.");
    state.Require(readFakeInstrumentationCounter("fileSizeCalls", fileSizeCallsAfterDirectorySize, L"after directory size"),
                  L"MTP fake enumerate: failed to read final fileSizeCalls.");
    state.Require(readFileCallsAfterDirectorySize == readFileCallsBeforeDirectorySize,
                  std::format(L"MTP fake enumerate: GetDirectorySize read file contents. before={} after={}.",
                              readFileCallsBeforeDirectorySize,
                              readFileCallsAfterDirectorySize));
    state.Require(fileSizeCallsAfterDirectorySize > fileSizeCallsBeforeDirectorySize,
                  std::format(L"MTP fake enumerate: GetDirectorySize did not use file-size metadata. before={} after={}.",
                              fileSizeCallsBeforeDirectorySize,
                              fileSizeCallsAfterDirectorySize));
    state.Require(ValidateDirectorySizeCallbackContract(state, dirOps.get(), std::wstring(kCameraPath), FILESYSTEM_FLAG_RECURSIVE),
                  L"MTP fake enumerate: directory size callback contract failed.");

    return state.failure.empty();
});

SelfTest::RunCase(options,
                  suite,
                  L"mtp_overwrite_byte_verify_level_matches_capability",
                  [&](SelfTest::CaseState& state) noexcept
{
    const auto widenAscii = [](std::string_view text) { return std::wstring(text.begin(), text.end()); };

    struct VerifyExpectation
    {
        std::string_view level;
        uint32_t expectedReadFileCalls = 0;
    };

    constexpr std::array<VerifyExpectation, 3> kExpectations = {{
        {.level = "transmitHash", .expectedReadFileCalls = 0},
        {.level = "deviceReread", .expectedReadFileCalls = 1},
        {.level = "sizeOnly", .expectedReadFileCalls = 0},
    }};

    const std::wstring guid = MakeGuidText();
    state.Require(! guid.empty(), L"MTP overwrite verify: failed to generate unique case name.");
    if (guid.empty())
    {
        return false;
    }

    for (const VerifyExpectation& expectation : kExpectations)
    {
        CreatedFileSystemInstance created;
        const std::string configuration = std::format(R"json({{"readOnly":false,"byteVerifyOnOverwrite":"{}"}})json", expectation.level);
        const HRESULT createHr          = TryCreateFakeMtpFileSystemInstance("{}", configuration, L"/", created);
        if (createHr == HRESULT_FROM_WIN32(ERROR_PROC_NOT_FOUND))
        {
            return state.Skip(L"MTP fake backend export is available only in debug plugin builds.");
        }
        state.Require(SUCCEEDED(createHr) && created.fileSystem,
                      std::format(L"MTP overwrite verify: create selftest instance for {} failed. hr=0x{:08X}",
                                  widenAscii(expectation.level),
                                  static_cast<unsigned long>(createHr)));
        if (FAILED(createHr) || ! created.fileSystem)
        {
            return false;
        }

        wil::com_ptr<IFileSystemIO> io;
        state.Require(CreateFileSystemIo(created.fileSystem, io),
                      std::format(L"MTP overwrite verify: {} missing IFileSystemIO.", widenAscii(expectation.level)));
        if (! io)
        {
            return false;
        }

        const char* capabilities = nullptr;
        const HRESULT capsHr     = created.fileSystem->GetCapabilities(&capabilities);
        state.Require(
            SUCCEEDED(capsHr) && capabilities != nullptr,
            std::format(L"MTP overwrite verify: {} GetCapabilities failed. hr=0x{:08X}", widenAscii(expectation.level), static_cast<unsigned long>(capsHr)));
        if (SUCCEEDED(capsHr) && capabilities)
        {
            const std::string expectedCapability = std::format(R"json("byteVerifyOnOverwrite": "{}")json", expectation.level);
            state.Require(std::string_view(capabilities).find(expectedCapability) != std::string_view::npos,
                          std::format(L"MTP overwrite verify: capabilities did not advertise {}.", widenAscii(expectedCapability)));
        }

        const std::wstring levelName(expectation.level.begin(), expectation.level.end());
        const std::wstring path = L"/Fake Phone [devpuid:fake-device]/Internal Storage/DCIM/Camera/verify-" + levelName + L"-" + guid + L".txt";
        auto cleanup            = wil::scope_exit([&]() noexcept
        { static_cast<void>(created.fileSystem->DeleteItem(path.c_str(), FILESYSTEM_FLAG_NONE, nullptr, nullptr, nullptr)); });

        constexpr std::string_view kInitialPayload = "initial overwrite verification payload";
        state.Require(WritePluginFileText(io.get(), path.c_str(), FILESYSTEM_FLAG_NONE, kInitialPayload, state, L"MTP overwrite verify initial write"),
                      std::format(L"MTP overwrite verify: {} initial write failed.", widenAscii(expectation.level)));
        if (! state.failure.empty())
        {
            return false;
        }

        constexpr std::string_view kOverwritePayload = "replacement payload for verification";
        state.Require(WritePluginFileText(io.get(), path.c_str(), FILESYSTEM_FLAG_ALLOW_OVERWRITE, kOverwritePayload, state, L"MTP overwrite verify overwrite"),
                      std::format(L"MTP overwrite verify: {} overwrite failed.", widenAscii(expectation.level)));
        if (! state.failure.empty())
        {
            return false;
        }

        const char* properties = nullptr;
        const HRESULT propsHr  = io->GetItemProperties(path.c_str(), &properties);
        state.Require(
            SUCCEEDED(propsHr) && properties != nullptr,
            std::format(L"MTP overwrite verify: {} GetItemProperties failed. hr=0x{:08X}", widenAscii(expectation.level), static_cast<unsigned long>(propsHr)));
        if (SUCCEEDED(propsHr) && properties)
        {
            const std::string_view props(properties);
            const std::string expectedWrites = R"json("writeFileCalls":2)json";
            const std::string expectedSizes  = R"json("fileSizeCalls":1)json";
            const std::string expectedReads  = std::format(R"json("readFileCalls":{})json", expectation.expectedReadFileCalls);

            state.Require(props.find(expectedWrites) != std::string_view::npos,
                          std::format(L"MTP overwrite verify: {} did not perform exactly two backend writes.", widenAscii(expectation.level)));
            state.Require(props.find(expectedSizes) != std::string_view::npos,
                          std::format(L"MTP overwrite verify: {} did not perform exactly one post-commit size check.", widenAscii(expectation.level)));
            state.Require(props.find(expectedReads) != std::string_view::npos,
                          std::format(L"MTP overwrite verify: {} expected token {}.", widenAscii(expectation.level), widenAscii(expectedReads)));
        }

        std::string readBack;
        state.Require(ReadPluginFileText(io.get(), path.c_str(), readBack, state, L"MTP overwrite verify readback"),
                      std::format(L"MTP overwrite verify: {} readback failed.", widenAscii(expectation.level)));
        state.Require(readBack == kOverwritePayload, std::format(L"MTP overwrite verify: {} readback contents did not match.", widenAscii(expectation.level)));
    }

    return state.failure.empty();
});

SelfTest::RunCase(options,
                  suite,
                  L"mtp_writer_overwrite_uses_temp_puid_swap",
                  [&](SelfTest::CaseState& state) noexcept
{
    CreatedFileSystemInstance created;
    const HRESULT createHr = TryCreateFakeMtpFileSystemInstance("{}", R"json({"readOnly":false})json", L"/", created);
    if (createHr == HRESULT_FROM_WIN32(ERROR_PROC_NOT_FOUND))
    {
        return state.Skip(L"MTP fake backend export is available only in debug plugin builds.");
    }
    state.Require(SUCCEEDED(createHr) && created.fileSystem,
                  std::format(L"MTP writer overwrite swap: create selftest instance failed. hr=0x{:08X}", static_cast<unsigned long>(createHr)));
    if (FAILED(createHr) || ! created.fileSystem)
    {
        return false;
    }

    wil::com_ptr<IFileSystemIO> io;
    state.Require(CreateFileSystemIo(created.fileSystem, io), L"MTP writer overwrite swap: missing IFileSystemIO.");
    if (! io)
    {
        return false;
    }

    const auto widenAscii        = [](std::string_view text) { return std::wstring(text.begin(), text.end()); };
    const auto extractJsonString = [](std::string_view json, std::string_view key) noexcept -> std::string
    {
        const std::string token = std::format(R"json("{}":")json", key);
        const size_t start      = json.find(token);
        if (start == std::string_view::npos)
        {
            return {};
        }

        const size_t valueStart = start + token.size();
        const size_t valueEnd   = json.find('"', valueStart);
        if (valueEnd == std::string_view::npos)
        {
            return {};
        }

        return std::string(json.substr(valueStart, valueEnd - valueStart));
    };
    const auto readPersistentId = [&](const std::wstring& path, std::string& persistentId, std::wstring_view label) noexcept
    {
        persistentId.clear();
        const char* properties = nullptr;
        const HRESULT propsHr  = io->GetItemProperties(path.c_str(), &properties);
        state.Require(SUCCEEDED(propsHr) && properties != nullptr,
                      std::format(L"MTP writer overwrite swap: {} GetItemProperties failed. hr=0x{:08X}", label, static_cast<unsigned long>(propsHr)));
        if (FAILED(propsHr) || properties == nullptr)
        {
            return;
        }

        persistentId = extractJsonString(properties, "persistentId");
        state.Require(! persistentId.empty(), std::format(L"MTP writer overwrite swap: {} did not expose a persistent id.", label));
    };

    const std::wstring guid = MakeGuidText();
    state.Require(! guid.empty(), L"MTP writer overwrite swap: failed to generate unique file name.");
    if (guid.empty())
    {
        return false;
    }

    const std::wstring baseFolder = L"/Fake Phone [devpuid:fake-device]/Internal Storage/DCIM/Camera";
    const std::wstring fileName   = L"safe-overwrite-" + guid + L".txt";
    const std::wstring path       = baseFolder + L"/" + fileName;
    auto cleanup =
        wil::scope_exit([&]() noexcept { static_cast<void>(created.fileSystem->DeleteItem(path.c_str(), FILESYSTEM_FLAG_NONE, nullptr, nullptr, nullptr)); });

    constexpr std::string_view kInitialPayload = "initial payload before safe overwrite";
    state.Require(WritePluginFileText(io.get(), path.c_str(), FILESYSTEM_FLAG_NONE, kInitialPayload, state, L"MTP writer overwrite swap initial write"),
                  L"MTP writer overwrite swap: initial write failed.");
    if (! state.failure.empty())
    {
        return false;
    }

    std::string initialPersistentId;
    readPersistentId(path, initialPersistentId, L"initial");
    if (! state.failure.empty())
    {
        return false;
    }

    constexpr std::string_view kOverwritePayload = "replacement payload after temp swap";
    state.Require(
        WritePluginFileText(io.get(), path.c_str(), FILESYSTEM_FLAG_ALLOW_OVERWRITE, kOverwritePayload, state, L"MTP writer overwrite swap overwrite"),
        L"MTP writer overwrite swap: overwrite write failed.");
    if (! state.failure.empty())
    {
        return false;
    }

    std::string finalPersistentId;
    readPersistentId(path, finalPersistentId, L"final");
    state.Require(! finalPersistentId.empty() && finalPersistentId != initialPersistentId,
                  L"MTP writer overwrite swap: final path did not take the temp object's persistent id.");

    std::string readBack;
    state.Require(ReadPluginFileText(io.get(), path.c_str(), readBack, state, L"MTP writer overwrite swap readback"),
                  L"MTP writer overwrite swap: readback failed.");
    state.Require(readBack == kOverwritePayload, L"MTP writer overwrite swap: final contents did not match replacement payload.");

    const auto cameraEntries    = SnapshotDirectoryEntries(created.fileSystem, baseFolder.c_str(), state, L"MTP writer overwrite swap camera");
    unsigned int finalNameCount = 0;
    for (const DirectoryEntrySnapshot& entry : cameraEntries)
    {
        if (entry.name == fileName)
        {
            ++finalNameCount;
        }
        state.Require(entry.name.find(L".rs-mtp-overwrite-") == std::wstring::npos,
                      std::format(L"MTP writer overwrite swap: leaked temp entry {}.", entry.name));
    }

    state.Require(finalNameCount == 1u,
                  std::format(L"MTP writer overwrite swap: expected one final entry, got {}. initialPuid={} finalPuid={}",
                              finalNameCount,
                              widenAscii(initialPersistentId),
                              widenAscii(finalPersistentId)));
    return state.failure.empty();
});

SelfTest::RunCase(options,
                  suite,
                  L"mtp_overwrite_temp_upload_failure_keeps_original_and_allows_retry",
                  [&](SelfTest::CaseState& state) noexcept
{
    CreatedFileSystemInstance created;
    const HRESULT createHr =
        TryCreateFakeMtpFileSystemInstance(R"json({"writeFileFailOncePathContains":".rs-mtp-overwrite-"})json", R"json({"readOnly":false})json", L"/", created);
    if (createHr == HRESULT_FROM_WIN32(ERROR_PROC_NOT_FOUND))
    {
        return state.Skip(L"MTP fake backend export is available only in debug plugin builds.");
    }
    state.Require(SUCCEEDED(createHr) && created.fileSystem,
                  std::format(L"MTP temp-upload-failure overwrite: create selftest instance failed. hr=0x{:08X}", static_cast<unsigned long>(createHr)));
    if (FAILED(createHr) || ! created.fileSystem)
    {
        return false;
    }

    wil::com_ptr<IFileSystemIO> io;
    state.Require(CreateFileSystemIo(created.fileSystem, io), L"MTP temp-upload-failure overwrite: missing IFileSystemIO.");
    if (! io)
    {
        return false;
    }

    const auto readWriteCalls = [&](const std::wstring& path, uint64_t& writeCalls, std::wstring_view label) noexcept
    {
        writeCalls             = 0;
        const char* properties = nullptr;
        const HRESULT propsHr  = io->GetItemProperties(path.c_str(), &properties);
        state.Require(SUCCEEDED(propsHr) && properties != nullptr,
                      std::format(L"MTP temp-upload-failure overwrite: {} GetItemProperties failed. hr=0x{:08X}", label, static_cast<unsigned long>(propsHr)));
        if (FAILED(propsHr) || properties == nullptr)
        {
            return;
        }

        const std::optional<uint64_t> parsed = extractJsonUInt(properties, "writeFileCalls");
        state.Require(parsed.has_value(), std::format(L"MTP temp-upload-failure overwrite: {} missing writeFileCalls.", label));
        if (parsed.has_value())
        {
            writeCalls = parsed.value();
        }
    };

    const std::wstring guid = MakeGuidText();
    state.Require(! guid.empty(), L"MTP temp-upload-failure overwrite: failed to generate unique file name.");
    if (guid.empty())
    {
        return false;
    }

    const std::wstring baseFolder = L"/Fake Phone [devpuid:fake-device]/Internal Storage/DCIM/Camera";
    const std::wstring fileName   = L"temp-upload-failure-" + guid + L".txt";
    const std::wstring path       = baseFolder + L"/" + fileName;
    auto cleanup =
        wil::scope_exit([&]() noexcept { static_cast<void>(created.fileSystem->DeleteItem(path.c_str(), FILESYSTEM_FLAG_NONE, nullptr, nullptr, nullptr)); });

    const auto requireSingleFinalAndNoTemp = [&](std::wstring_view label) noexcept
    {
        const auto entries      = SnapshotDirectoryEntries(created.fileSystem, baseFolder.c_str(), state, L"MTP temp-upload-failure camera");
        uint32_t finalNameCount = 0;
        for (const DirectoryEntrySnapshot& entry : entries)
        {
            if (entry.name == fileName)
            {
                ++finalNameCount;
            }
            state.Require(entry.name.find(L".rs-mtp-overwrite-") == std::wstring::npos,
                          std::format(L"MTP temp-upload-failure overwrite: leaked temp entry {} {}.", entry.name, label));
        }
        state.Require(finalNameCount == 1u, std::format(L"MTP temp-upload-failure overwrite: expected one final entry {}, got {}.", label, finalNameCount));
    };

    const auto commitOverwrite = [&](std::string_view payload, std::wstring_view label) noexcept -> HRESULT
    {
        wil::com_ptr<IFileWriter> writer;
        HRESULT hr = io->CreateFileWriter(path.c_str(), FILESYSTEM_FLAG_ALLOW_OVERWRITE, writer.put());
        state.Require(SUCCEEDED(hr) && writer,
                      std::format(L"MTP temp-upload-failure overwrite: {} CreateFileWriter failed. hr=0x{:08X}", label, static_cast<unsigned long>(hr)));
        if (FAILED(hr) || ! writer)
        {
            return hr;
        }

        const unsigned long expectedBytes = static_cast<unsigned long>(payload.size());
        unsigned long written             = 0;
        hr                                = writer->Write(payload.data(), expectedBytes, &written);
        state.Require(SUCCEEDED(hr) && written == expectedBytes,
                      std::format(L"MTP temp-upload-failure overwrite: {} Write failed. wrote={} expected={} hr=0x{:08X}",
                                  label,
                                  written,
                                  expectedBytes,
                                  static_cast<unsigned long>(hr)));
        if (FAILED(hr) || written != expectedBytes)
        {
            return hr;
        }

        return writer->Commit();
    };

    constexpr std::string_view kInitialPayload = "original payload before temp upload failure";
    state.Require(WritePluginFileText(io.get(), path.c_str(), FILESYSTEM_FLAG_NONE, kInitialPayload, state, L"MTP temp-upload-failure initial write"),
                  L"MTP temp-upload-failure overwrite: initial write failed.");
    if (! state.failure.empty())
    {
        return false;
    }

    uint64_t writesBeforeFailure = 0;
    readWriteCalls(path, writesBeforeFailure, L"before failure");
    state.Require(writesBeforeFailure == 1u, std::format(L"MTP temp-upload-failure overwrite: expected one initial write, got {}.", writesBeforeFailure));
    if (! state.failure.empty())
    {
        return false;
    }

    constexpr std::string_view kFailedPayload = "replacement that fails while uploading temp";
    const HRESULT failedOverwriteHr           = commitOverwrite(kFailedPayload, L"failing overwrite");
    state.Require(
        failedOverwriteHr == HRESULT_FROM_WIN32(ERROR_WRITE_FAULT),
        std::format(L"MTP temp-upload-failure overwrite: expected ERROR_WRITE_FAULT, got hr=0x{:08X}.", static_cast<unsigned long>(failedOverwriteHr)));
    if (! state.failure.empty())
    {
        return false;
    }

    std::string readBack;
    state.Require(ReadPluginFileText(io.get(), path.c_str(), readBack, state, L"MTP temp-upload-failure readback after failure"),
                  L"MTP temp-upload-failure overwrite: readback after failure failed.");
    state.Require(readBack == kInitialPayload, L"MTP temp-upload-failure overwrite: failed temp upload changed original contents.");
    requireSingleFinalAndNoTemp(L"after failed upload");

    uint64_t writesAfterFailure = 0;
    readWriteCalls(path, writesAfterFailure, L"after failure");
    state.Require(writesAfterFailure == writesBeforeFailure + 1u,
                  std::format(L"MTP temp-upload-failure overwrite: expected exactly one failed temp write. before={} after={}.",
                              writesBeforeFailure,
                              writesAfterFailure));
    if (! state.failure.empty())
    {
        return false;
    }

    constexpr std::string_view kRetryPayload = "replacement payload after temp upload retry";
    const HRESULT retryHr                    = commitOverwrite(kRetryPayload, L"retry overwrite");
    state.Require(SUCCEEDED(retryHr), std::format(L"MTP temp-upload-failure overwrite: retry failed. hr=0x{:08X}", static_cast<unsigned long>(retryHr)));
    if (! state.failure.empty())
    {
        return false;
    }

    readBack.clear();
    state.Require(ReadPluginFileText(io.get(), path.c_str(), readBack, state, L"MTP temp-upload-failure retry readback"),
                  L"MTP temp-upload-failure overwrite: retry readback failed.");
    state.Require(readBack == kRetryPayload, L"MTP temp-upload-failure overwrite: retry contents did not match.");
    requireSingleFinalAndNoTemp(L"after retry");

    return state.failure.empty();
});

SelfTest::RunCase(options,
                  suite,
                  L"mtp_overwrite_empty_temp_puid_keeps_original_and_blocks_later_upload",
                  [&](SelfTest::CaseState& state) noexcept
{
    CreatedFileSystemInstance created;
    const HRESULT createHr =
        TryCreateFakeMtpFileSystemInstance(R"json({"omitPersistentIdForCreatedFiles":true})json", R"json({"readOnly":false})json", L"/", created);
    if (createHr == HRESULT_FROM_WIN32(ERROR_PROC_NOT_FOUND))
    {
        return state.Skip(L"MTP fake backend export is available only in debug plugin builds.");
    }
    state.Require(SUCCEEDED(createHr) && created.fileSystem,
                  std::format(L"MTP empty-temp-PUID overwrite: create selftest instance failed. hr=0x{:08X}", static_cast<unsigned long>(createHr)));
    if (FAILED(createHr) || ! created.fileSystem)
    {
        return false;
    }

    wil::com_ptr<IFileSystemIO> io;
    state.Require(CreateFileSystemIo(created.fileSystem, io), L"MTP empty-temp-PUID overwrite: missing IFileSystemIO.");
    if (! io)
    {
        return false;
    }

    const std::wstring guid = MakeGuidText();
    state.Require(! guid.empty(), L"MTP empty-temp-PUID overwrite: failed to generate unique file name.");
    if (guid.empty())
    {
        return false;
    }

    const std::wstring baseFolder = L"/Fake Phone [devpuid:fake-device]/Internal Storage/DCIM/Camera";
    const std::wstring fileName   = L"empty-temp-puid-" + guid + L".txt";
    const std::wstring path       = baseFolder + L"/" + fileName;
    auto cleanup =
        wil::scope_exit([&]() noexcept { static_cast<void>(created.fileSystem->DeleteItem(path.c_str(), FILESYSTEM_FLAG_NONE, nullptr, nullptr, nullptr)); });
    const auto readWriteCalls = [&](uint64_t& writeCalls, std::wstring_view label) noexcept
    {
        writeCalls             = 0;
        const char* properties = nullptr;
        const HRESULT propsHr  = io->GetItemProperties(path.c_str(), &properties);
        state.Require(SUCCEEDED(propsHr) && properties != nullptr,
                      std::format(L"MTP empty-temp-PUID overwrite: {} GetItemProperties failed. hr=0x{:08X}", label, static_cast<unsigned long>(propsHr)));
        if (FAILED(propsHr) || properties == nullptr)
        {
            return;
        }

        const std::optional<uint64_t> parsed = extractJsonUInt(properties, "writeFileCalls");
        state.Require(parsed.has_value(), std::format(L"MTP empty-temp-PUID overwrite: {} missing writeFileCalls in properties.", label));
        if (parsed.has_value())
        {
            writeCalls = parsed.value();
        }
    };

    const auto commitOverwrite = [&](std::string_view payload, std::wstring_view label) noexcept -> HRESULT
    {
        wil::com_ptr<IFileWriter> writer;
        HRESULT hr = io->CreateFileWriter(path.c_str(), FILESYSTEM_FLAG_ALLOW_OVERWRITE, writer.put());
        state.Require(SUCCEEDED(hr) && writer,
                      std::format(L"MTP empty-temp-PUID overwrite: {} CreateFileWriter failed. hr=0x{:08X}", label, static_cast<unsigned long>(hr)));
        if (FAILED(hr) || ! writer)
        {
            return hr;
        }

        const unsigned long expectedBytes = static_cast<unsigned long>(payload.size());
        unsigned long written             = 0;
        hr                                = writer->Write(payload.data(), expectedBytes, &written);
        state.Require(SUCCEEDED(hr) && written == expectedBytes,
                      std::format(L"MTP empty-temp-PUID overwrite: {} Write failed. wrote={} expected={} hr=0x{:08X}",
                                  label,
                                  written,
                                  expectedBytes,
                                  static_cast<unsigned long>(hr)));
        if (FAILED(hr) || written != expectedBytes)
        {
            return hr;
        }

        return writer->Commit();
    };

    constexpr std::string_view kInitialPayload = "original payload on PUID-less fake device";
    state.Require(WritePluginFileText(io.get(), path.c_str(), FILESYSTEM_FLAG_NONE, kInitialPayload, state, L"MTP empty-temp-PUID initial write"),
                  L"MTP empty-temp-PUID overwrite: initial write failed.");
    if (! state.failure.empty())
    {
        return false;
    }

    constexpr std::string_view kReplacementPayload = "replacement that must not delete original";
    const HRESULT firstOverwriteHr                 = commitOverwrite(kReplacementPayload, L"first overwrite");
    state.Require(firstOverwriteHr == HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED),
                  std::format(L"MTP empty-temp-PUID overwrite: first overwrite expected ERROR_NOT_SUPPORTED, got hr=0x{:08X}.",
                              static_cast<unsigned long>(firstOverwriteHr)));
    if (! state.failure.empty())
    {
        return false;
    }

    std::string readBack;
    state.Require(ReadPluginFileText(io.get(), path.c_str(), readBack, state, L"MTP empty-temp-PUID first readback"),
                  L"MTP empty-temp-PUID overwrite: readback after first failure failed.");
    state.Require(readBack == kInitialPayload, L"MTP empty-temp-PUID overwrite: first failure did not preserve original contents.");

    uint64_t writesAfterFirstFailure = 0;
    readWriteCalls(writesAfterFirstFailure, L"after first failure");
    state.Require(writesAfterFirstFailure == 2u,
                  std::format(L"MTP empty-temp-PUID overwrite: expected 2 writes after first failure, got {}.", writesAfterFirstFailure));

    const auto entriesAfterFirstFailure =
        SnapshotDirectoryEntries(created.fileSystem, baseFolder.c_str(), state, L"MTP empty-temp-PUID camera after first failure");
    uint32_t finalNameCount = 0;
    for (const DirectoryEntrySnapshot& entry : entriesAfterFirstFailure)
    {
        if (entry.name == fileName)
        {
            ++finalNameCount;
        }
        state.Require(entry.name.find(L".rs-mtp-overwrite-") == std::wstring::npos,
                      std::format(L"MTP empty-temp-PUID overwrite: leaked temp entry {} after first failure.", entry.name));
    }
    state.Require(finalNameCount == 1u, std::format(L"MTP empty-temp-PUID overwrite: expected one final entry after first failure, got {}.", finalNameCount));
    if (! state.failure.empty())
    {
        return false;
    }

    constexpr std::string_view kSecondReplacementPayload = "second replacement must be blocked before upload";
    const HRESULT secondOverwriteHr                      = commitOverwrite(kSecondReplacementPayload, L"second overwrite");
    state.Require(secondOverwriteHr == HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED),
                  std::format(L"MTP empty-temp-PUID overwrite: second overwrite expected ERROR_NOT_SUPPORTED, got hr=0x{:08X}.",
                              static_cast<unsigned long>(secondOverwriteHr)));
    if (! state.failure.empty())
    {
        return false;
    }

    uint64_t writesAfterSecondFailure = 0;
    readWriteCalls(writesAfterSecondFailure, L"after second failure");
    state.Require(writesAfterSecondFailure == writesAfterFirstFailure,
                  std::format(L"MTP empty-temp-PUID overwrite: policy did not block second upload. before={} after={}.",
                              writesAfterFirstFailure,
                              writesAfterSecondFailure));

    readBack.clear();
    state.Require(ReadPluginFileText(io.get(), path.c_str(), readBack, state, L"MTP empty-temp-PUID second readback"),
                  L"MTP empty-temp-PUID overwrite: readback after second failure failed.");
    state.Require(readBack == kInitialPayload, L"MTP empty-temp-PUID overwrite: second failure did not preserve original contents.");

    const auto entriesAfterSecondFailure =
        SnapshotDirectoryEntries(created.fileSystem, baseFolder.c_str(), state, L"MTP empty-temp-PUID camera after second failure");
    finalNameCount = 0;
    for (const DirectoryEntrySnapshot& entry : entriesAfterSecondFailure)
    {
        if (entry.name == fileName)
        {
            ++finalNameCount;
        }
        state.Require(entry.name.find(L".rs-mtp-overwrite-") == std::wstring::npos,
                      std::format(L"MTP empty-temp-PUID overwrite: leaked temp entry {} after second failure.", entry.name));
    }
    state.Require(finalNameCount == 1u, std::format(L"MTP empty-temp-PUID overwrite: expected one final entry after second failure, got {}.", finalNameCount));

    return state.failure.empty();
});

SelfTest::RunCase(options,
                  suite,
                  L"mtp_overwrite_delete_original_failure_keeps_original_and_allows_retry",
                  [&](SelfTest::CaseState& state) noexcept
{
    const std::wstring guid = MakeGuidText();
    state.Require(! guid.empty(), L"MTP delete-original-failure overwrite: failed to generate unique file name.");
    if (guid.empty())
    {
        return false;
    }

    const std::wstring baseFolder = L"/Fake Phone [devpuid:fake-device]/Internal Storage/DCIM/Camera";
    const std::wstring fileName   = L"delete-original-failure-" + guid + L".txt";
    const std::wstring path       = baseFolder + L"/" + fileName;
    const std::string fakeOptions = std::format(R"json({{"deleteItemFailOncePath":"{}"}})json", narrowAscii(path));

    CreatedFileSystemInstance created;
    const HRESULT createHr = TryCreateFakeMtpFileSystemInstance(fakeOptions, R"json({"readOnly":false})json", L"/", created);
    if (createHr == HRESULT_FROM_WIN32(ERROR_PROC_NOT_FOUND))
    {
        return state.Skip(L"MTP fake backend export is available only in debug plugin builds.");
    }
    state.Require(SUCCEEDED(createHr) && created.fileSystem,
                  std::format(L"MTP delete-original-failure overwrite: create selftest instance failed. hr=0x{:08X}", static_cast<unsigned long>(createHr)));
    if (FAILED(createHr) || ! created.fileSystem)
    {
        return false;
    }

    wil::com_ptr<IFileSystemIO> io;
    state.Require(CreateFileSystemIo(created.fileSystem, io), L"MTP delete-original-failure overwrite: missing IFileSystemIO.");
    if (! io)
    {
        return false;
    }

    auto cleanup =
        wil::scope_exit([&]() noexcept { static_cast<void>(created.fileSystem->DeleteItem(path.c_str(), FILESYSTEM_FLAG_NONE, nullptr, nullptr, nullptr)); });

    const auto commitOverwrite = [&](std::string_view payload, std::wstring_view label) noexcept -> HRESULT
    {
        wil::com_ptr<IFileWriter> writer;
        HRESULT hr = io->CreateFileWriter(path.c_str(), FILESYSTEM_FLAG_ALLOW_OVERWRITE, writer.put());
        state.Require(SUCCEEDED(hr) && writer,
                      std::format(L"MTP delete-original-failure overwrite: {} CreateFileWriter failed. hr=0x{:08X}", label, static_cast<unsigned long>(hr)));
        if (FAILED(hr) || ! writer)
        {
            return hr;
        }

        const unsigned long expectedBytes = static_cast<unsigned long>(payload.size());
        unsigned long written             = 0;
        hr                                = writer->Write(payload.data(), expectedBytes, &written);
        state.Require(SUCCEEDED(hr) && written == expectedBytes,
                      std::format(L"MTP delete-original-failure overwrite: {} Write failed. wrote={} expected={} hr=0x{:08X}",
                                  label,
                                  written,
                                  expectedBytes,
                                  static_cast<unsigned long>(hr)));
        if (FAILED(hr) || written != expectedBytes)
        {
            return hr;
        }

        return writer->Commit();
    };

    const auto requireSingleFinalAndNoTemp = [&](std::wstring_view label) noexcept
    {
        const auto entries      = SnapshotDirectoryEntries(created.fileSystem, baseFolder.c_str(), state, label);
        uint32_t finalNameCount = 0;
        for (const DirectoryEntrySnapshot& entry : entries)
        {
            if (entry.name == fileName)
            {
                ++finalNameCount;
            }
            state.Require(entry.name.find(L".rs-mtp-overwrite-") == std::wstring::npos,
                          std::format(L"MTP delete-original-failure overwrite: leaked temp entry {} during {}.", entry.name, label));
        }
        state.Require(finalNameCount == 1u,
                      std::format(L"MTP delete-original-failure overwrite: expected one final entry during {}, got {}.", label, finalNameCount));
    };

    constexpr std::string_view kInitialPayload = "original payload before delete-original failure";
    state.Require(WritePluginFileText(io.get(), path.c_str(), FILESYSTEM_FLAG_NONE, kInitialPayload, state, L"MTP delete-original-failure initial write"),
                  L"MTP delete-original-failure overwrite: initial write failed.");
    if (! state.failure.empty())
    {
        return false;
    }

    constexpr std::string_view kFailedReplacementPayload = "replacement blocked by delete-original failure";
    const HRESULT firstOverwriteHr                       = commitOverwrite(kFailedReplacementPayload, L"first overwrite");
    state.Require(firstOverwriteHr == HRESULT_FROM_WIN32(ERROR_ACCESS_DENIED),
                  std::format(L"MTP delete-original-failure overwrite: first overwrite expected ERROR_ACCESS_DENIED, got hr=0x{:08X}.",
                              static_cast<unsigned long>(firstOverwriteHr)));
    if (! state.failure.empty())
    {
        return false;
    }

    std::string readBack;
    state.Require(ReadPluginFileText(io.get(), path.c_str(), readBack, state, L"MTP delete-original-failure first readback"),
                  L"MTP delete-original-failure overwrite: readback after first failure failed.");
    state.Require(readBack == kInitialPayload, L"MTP delete-original-failure overwrite: first failure did not preserve original contents.");
    requireSingleFinalAndNoTemp(L"after first failure");
    if (! state.failure.empty())
    {
        return false;
    }

    constexpr std::string_view kRetryPayload = "replacement after delete-original retry succeeds";
    state.Require(
        WritePluginFileText(io.get(), path.c_str(), FILESYSTEM_FLAG_ALLOW_OVERWRITE, kRetryPayload, state, L"MTP delete-original-failure retry write"),
        L"MTP delete-original-failure overwrite: retry write failed.");
    if (! state.failure.empty())
    {
        return false;
    }

    readBack.clear();
    state.Require(ReadPluginFileText(io.get(), path.c_str(), readBack, state, L"MTP delete-original-failure retry readback"),
                  L"MTP delete-original-failure overwrite: retry readback failed.");
    state.Require(readBack == kRetryPayload, L"MTP delete-original-failure overwrite: retry contents did not match.");
    requireSingleFinalAndNoTemp(L"after retry");

    return state.failure.empty();
});

SelfTest::RunCase(options,
                  suite,
                  L"mtp_overwrite_journal_write_failure_aborts_before_upload",
                  [&](SelfTest::CaseState& state) noexcept
{
    enum class Operation : uint8_t
    {
        Writer,
        Copy,
        Move,
    };

    struct Scenario
    {
        Operation operation;
        std::wstring_view label;
    };

    constexpr std::array<Scenario, 3> kScenarios = {{
        {.operation = Operation::Writer, .label = L"writer"},
        {.operation = Operation::Copy, .label = L"copy"},
        {.operation = Operation::Move, .label = L"move"},
    }};

    const auto readCounters = [&](IFileSystemIO* io, const wchar_t* path, uint64_t& writeCalls, uint64_t& copyItemCalls, std::wstring_view label) noexcept
    {
        writeCalls             = 0;
        copyItemCalls          = 0;
        const char* properties = nullptr;
        const HRESULT propsHr  = io->GetItemProperties(path, &properties);
        state.Require(SUCCEEDED(propsHr) && properties != nullptr,
                      std::format(L"MTP journal failure: {} GetItemProperties failed. hr=0x{:08X}", label, static_cast<unsigned long>(propsHr)));
        if (FAILED(propsHr) || properties == nullptr)
        {
            return;
        }

        const std::optional<uint64_t> writes = extractJsonUInt(properties, "writeFileCalls");
        const std::optional<uint64_t> copies = extractJsonUInt(properties, "copyItemCalls");
        state.Require(writes.has_value(), std::format(L"MTP journal failure: {} missing writeFileCalls.", label));
        state.Require(copies.has_value(), std::format(L"MTP journal failure: {} missing copyItemCalls.", label));
        writeCalls    = writes.value_or(0u);
        copyItemCalls = copies.value_or(0u);
    };

    const auto commitWriterOverwrite = [&](IFileSystemIO* io, const std::wstring& path, std::wstring_view label) noexcept -> HRESULT
    {
        wil::com_ptr<IFileWriter> writer;
        HRESULT hr = io->CreateFileWriter(path.c_str(), FILESYSTEM_FLAG_ALLOW_OVERWRITE, writer.put());
        state.Require(SUCCEEDED(hr) && writer,
                      std::format(L"MTP journal failure: {} CreateFileWriter failed. hr=0x{:08X}", label, static_cast<unsigned long>(hr)));
        if (FAILED(hr) || ! writer)
        {
            return hr;
        }

        constexpr std::string_view kReplacement = "replacement blocked before upload by journal failure";
        unsigned long written                   = 0;
        hr                                      = writer->Write(kReplacement.data(), static_cast<unsigned long>(kReplacement.size()), &written);
        state.Require(SUCCEEDED(hr) && written == kReplacement.size(),
                      std::format(L"MTP journal failure: {} Write failed. wrote={} expected={} hr=0x{:08X}",
                                  label,
                                  written,
                                  static_cast<unsigned long>(kReplacement.size()),
                                  static_cast<unsigned long>(hr)));
        if (FAILED(hr) || written != kReplacement.size())
        {
            return hr;
        }

        return writer->Commit();
    };

    const std::wstring guid = MakeGuidText();
    state.Require(! guid.empty(), L"MTP journal failure: failed to generate unique case name.");
    if (guid.empty())
    {
        return false;
    }

    const std::wstring baseFolder = L"/Fake Phone [devpuid:fake-device]/Internal Storage/DCIM/Camera";
    for (const Scenario& scenario : kScenarios)
    {
        CreatedFileSystemInstance created;
        const HRESULT createHr = TryCreateFakeMtpFileSystemInstance("{}", R"json({"readOnly":false,"failOverwriteJournalWrites":true})json", L"/", created);
        if (createHr == HRESULT_FROM_WIN32(ERROR_PROC_NOT_FOUND))
        {
            return state.Skip(L"MTP fake backend export is available only in debug plugin builds.");
        }
        state.Require(
            SUCCEEDED(createHr) && created.fileSystem,
            std::format(L"MTP journal failure: create selftest instance for {} failed. hr=0x{:08X}", scenario.label, static_cast<unsigned long>(createHr)));
        if (FAILED(createHr) || ! created.fileSystem)
        {
            return false;
        }

        wil::com_ptr<IFileSystemIO> io;
        state.Require(CreateFileSystemIo(created.fileSystem, io), L"MTP journal failure: missing IFileSystemIO.");
        if (! io)
        {
            return false;
        }

        const std::wstring prefix          = std::format(L"journal-failure-{}-{}", scenario.label, guid);
        const std::wstring sourceLeaf      = prefix + L"-source.txt";
        const std::wstring destinationLeaf = prefix + L"-dest.txt";
        const std::wstring sourcePath      = baseFolder + L"/" + sourceLeaf;
        const std::wstring destinationPath = baseFolder + L"/" + destinationLeaf;
        auto cleanup                       = wil::scope_exit([&]() noexcept
        {
            static_cast<void>(created.fileSystem->DeleteItem(sourcePath.c_str(), FILESYSTEM_FLAG_NONE, nullptr, nullptr, nullptr));
            static_cast<void>(created.fileSystem->DeleteItem(destinationPath.c_str(), FILESYSTEM_FLAG_NONE, nullptr, nullptr, nullptr));
        });

        constexpr std::string_view kSourcePayload = "source payload that must not be uploaded";
        constexpr std::string_view kOldPayload    = "old destination payload preserved by journal failure";
        if (scenario.operation != Operation::Writer)
        {
            state.Require(WritePluginFileText(io.get(), sourcePath.c_str(), FILESYSTEM_FLAG_NONE, kSourcePayload, state, L"MTP journal failure source write"),
                          std::format(L"MTP journal failure: {} source write failed.", scenario.label));
        }
        state.Require(
            WritePluginFileText(io.get(), destinationPath.c_str(), FILESYSTEM_FLAG_NONE, kOldPayload, state, L"MTP journal failure destination write"),
            std::format(L"MTP journal failure: {} destination write failed.", scenario.label));
        if (! state.failure.empty())
        {
            return false;
        }

        uint64_t writesBefore = 0;
        uint64_t copiesBefore = 0;
        readCounters(io.get(), destinationPath.c_str(), writesBefore, copiesBefore, std::format(L"{} before", scenario.label));
        if (! state.failure.empty())
        {
            return false;
        }

        HRESULT operationHr = E_FAIL;
        switch (scenario.operation)
        {
            case Operation::Writer: operationHr = commitWriterOverwrite(io.get(), destinationPath, scenario.label); break;
            case Operation::Copy:
                operationHr =
                    created.fileSystem->CopyItem(sourcePath.c_str(), destinationPath.c_str(), FILESYSTEM_FLAG_ALLOW_OVERWRITE, nullptr, nullptr, nullptr);
                break;
            case Operation::Move:
                operationHr =
                    created.fileSystem->MoveItem(sourcePath.c_str(), destinationPath.c_str(), FILESYSTEM_FLAG_ALLOW_OVERWRITE, nullptr, nullptr, nullptr);
                break;
        }

        state.Require(
            operationHr == HRESULT_FROM_WIN32(ERROR_ACCESS_DENIED),
            std::format(L"MTP journal failure: {} expected ERROR_ACCESS_DENIED, got hr=0x{:08X}.", scenario.label, static_cast<unsigned long>(operationHr)));
        if (! state.failure.empty())
        {
            return false;
        }

        uint64_t writesAfter = 0;
        uint64_t copiesAfter = 0;
        readCounters(io.get(), destinationPath.c_str(), writesAfter, copiesAfter, std::format(L"{} after", scenario.label));
        state.Require(writesAfter == writesBefore,
                      std::format(L"MTP journal failure: {} touched backend writes. before={} after={}.", scenario.label, writesBefore, writesAfter));
        state.Require(copiesAfter == copiesBefore,
                      std::format(L"MTP journal failure: {} touched backend copies. before={} after={}.", scenario.label, copiesBefore, copiesAfter));

        std::string readBack;
        state.Require(ReadPluginFileText(io.get(), destinationPath.c_str(), readBack, state, L"MTP journal failure destination readback"),
                      std::format(L"MTP journal failure: {} destination readback failed.", scenario.label));
        state.Require(readBack == kOldPayload, std::format(L"MTP journal failure: {} did not preserve destination contents.", scenario.label));

        if (scenario.operation != Operation::Writer)
        {
            readBack.clear();
            state.Require(ReadPluginFileText(io.get(), sourcePath.c_str(), readBack, state, L"MTP journal failure source readback"),
                          std::format(L"MTP journal failure: {} source readback failed.", scenario.label));
            state.Require(readBack == kSourcePayload, std::format(L"MTP journal failure: {} did not preserve source contents.", scenario.label));
        }

        const auto entries      = SnapshotDirectoryEntries(created.fileSystem, baseFolder.c_str(), state, L"MTP journal failure camera");
        uint32_t finalNameCount = 0;
        for (const DirectoryEntrySnapshot& entry : entries)
        {
            if (entry.name == destinationLeaf)
            {
                ++finalNameCount;
            }
            state.Require(entry.name.find(L".rs-mtp-overwrite-") == std::wstring::npos,
                          std::format(L"MTP journal failure: leaked temp entry {} for {}.", entry.name, scenario.label));
        }
        state.Require(finalNameCount == 1u, std::format(L"MTP journal failure: {} expected one destination entry, got {}.", scenario.label, finalNameCount));

        if (! state.failure.empty())
        {
            return false;
        }
    }

    return state.failure.empty();
});

SelfTest::RunCase(options,
                  suite,
                  L"mtp_overwrite_journal_recovers_rename_temp_failure",
                  [&](SelfTest::CaseState& state) noexcept
{
    const std::wstring guid = MakeGuidText();
    state.Require(! guid.empty(), L"MTP journal rename recovery: failed to generate unique file name.");
    if (guid.empty())
    {
        return false;
    }

    const std::wstring baseFolder   = L"/Fake Phone [devpuid:fake-device]/Internal Storage/DCIM/Camera";
    const std::wstring fileName     = L"journal-rename-recovery-" + guid + L".txt";
    const std::wstring path         = baseFolder + L"/" + fileName;
    const std::string fakeOptions   = std::format(R"json({{"renameItemFailOnceDestinationPath":"{}"}})json", narrowAscii(path));
    const std::string configuration = std::format(R"json({{"readOnly":false,"host":"journal-rename-recovery-{}"}})json", narrowAscii(guid));

    CreatedFileSystemInstance created;
    const HRESULT createHr = TryCreateFakeMtpFileSystemInstance(fakeOptions, configuration, L"/", created);
    if (createHr == HRESULT_FROM_WIN32(ERROR_PROC_NOT_FOUND))
    {
        return state.Skip(L"MTP fake backend export is available only in debug plugin builds.");
    }
    state.Require(SUCCEEDED(createHr) && created.fileSystem,
                  std::format(L"MTP journal rename recovery: create selftest instance failed. hr=0x{:08X}", static_cast<unsigned long>(createHr)));
    if (FAILED(createHr) || ! created.fileSystem)
    {
        return false;
    }

    wil::com_ptr<IFileSystemIO> io;
    state.Require(CreateFileSystemIo(created.fileSystem, io), L"MTP journal rename recovery: missing IFileSystemIO.");
    if (! io)
    {
        return false;
    }

    auto cleanup =
        wil::scope_exit([&]() noexcept { static_cast<void>(created.fileSystem->DeleteItem(path.c_str(), FILESYSTEM_FLAG_NONE, nullptr, nullptr, nullptr)); });

    const auto commitOverwrite = [&](std::string_view payload, std::wstring_view label) noexcept -> HRESULT
    {
        wil::com_ptr<IFileWriter> writer;
        HRESULT hr = io->CreateFileWriter(path.c_str(), FILESYSTEM_FLAG_ALLOW_OVERWRITE, writer.put());
        state.Require(SUCCEEDED(hr) && writer,
                      std::format(L"MTP journal rename recovery: {} CreateFileWriter failed. hr=0x{:08X}", label, static_cast<unsigned long>(hr)));
        if (FAILED(hr) || ! writer)
        {
            return hr;
        }

        const unsigned long expectedBytes = static_cast<unsigned long>(payload.size());
        unsigned long written             = 0;
        hr                                = writer->Write(payload.data(), expectedBytes, &written);
        state.Require(SUCCEEDED(hr) && written == expectedBytes,
                      std::format(L"MTP journal rename recovery: {} Write failed. wrote={} expected={} hr=0x{:08X}",
                                  label,
                                  written,
                                  expectedBytes,
                                  static_cast<unsigned long>(hr)));
        if (FAILED(hr) || written != expectedBytes)
        {
            return hr;
        }

        return writer->Commit();
    };

    constexpr std::string_view kInitialPayload = "original payload before retained journal replay";
    state.Require(WritePluginFileText(io.get(), path.c_str(), FILESYSTEM_FLAG_NONE, kInitialPayload, state, L"MTP journal rename recovery initial write"),
                  L"MTP journal rename recovery: initial write failed.");
    if (! state.failure.empty())
    {
        return false;
    }

    constexpr std::string_view kReplacementPayload = "replacement recovered from retained temp journal";
    const HRESULT failedCommitHr                   = commitOverwrite(kReplacementPayload, L"rename-failing overwrite");
    state.Require(
        failedCommitHr == HRESULT_FROM_WIN32(ERROR_ACCESS_DENIED),
        std::format(L"MTP journal rename recovery: overwrite expected ERROR_ACCESS_DENIED, got hr=0x{:08X}.", static_cast<unsigned long>(failedCommitHr)));
    if (! state.failure.empty())
    {
        return false;
    }

    std::string readBack;
    state.Require(ReadPluginFileText(io.get(), path.c_str(), readBack, state, L"MTP journal rename recovery replay readback"),
                  L"MTP journal rename recovery: replay readback failed.");
    state.Require(readBack == kReplacementPayload, L"MTP journal rename recovery: retained temp was not replayed to final replacement contents.");

    const auto entries      = SnapshotDirectoryEntries(created.fileSystem, baseFolder.c_str(), state, L"MTP journal rename recovery camera");
    uint32_t finalNameCount = 0;
    for (const DirectoryEntrySnapshot& entry : entries)
    {
        if (entry.name == fileName)
        {
            ++finalNameCount;
        }
        state.Require(entry.name.find(L".rs-mtp-overwrite-") == std::wstring::npos,
                      std::format(L"MTP journal rename recovery: leaked temp entry {} after replay.", entry.name));
    }
    state.Require(finalNameCount == 1u, std::format(L"MTP journal rename recovery: expected one final entry after replay, got {}.", finalNameCount));

    return state.failure.empty();
});

SelfTest::RunCase(options,
                  suite,
                  L"mtp_overwrite_journal_replay_removes_temp_when_final_exists",
                  [&](SelfTest::CaseState& state) noexcept
{
    const std::wstring guid = MakeGuidText();
    state.Require(! guid.empty(), L"MTP journal orphan-temp cleanup: failed to generate unique file name.");
    if (guid.empty())
    {
        return false;
    }

    const std::wstring baseFolder   = L"/Fake Phone [devpuid:fake-device]/Internal Storage/DCIM/Camera";
    const std::wstring fileName     = L"journal-orphan-cleanup-" + guid + L".txt";
    const std::wstring tempName     = L".rs-mtp-overwrite-orphan-" + guid + L".tmp";
    const std::wstring finalPath    = baseFolder + L"/" + fileName;
    const std::wstring tempPath     = baseFolder + L"/" + tempName;
    const std::wstring host         = L"journal-orphan-cleanup-" + guid;
    const std::string configuration = std::format(R"json({{"readOnly":false,"host":"{}"}})json", narrowAscii(host));

    std::wstring previousLocalAppData;
    std::wstring localAppData;
    if (! AcquireMtpJournalLocalAppDataSandbox(
            state, L"mtp_journal_orphan_cleanup", L"MTP journal orphan-temp cleanup", localAppData, previousLocalAppData))
    {
        return false;
    }
    const auto restoreLocalAppData = wil::scope_exit([&]() noexcept { RestoreMtpJournalLocalAppDataSandbox(previousLocalAppData); });

    const std::wstring journalDirectory = localAppData + L"\\RedSalamander\\PluginState\\FileSystemMtp\\" + std::format(L"{:016X}", stableDeviceHash(host));
    const std::wstring journalPath      = journalDirectory + L"\\overwrite-journal.json";

    CreatedFileSystemInstance created;
    const HRESULT createHr = TryCreateFakeMtpFileSystemInstance("{}", configuration, L"/", created);
    if (createHr == HRESULT_FROM_WIN32(ERROR_PROC_NOT_FOUND))
    {
        return state.Skip(L"MTP fake backend export is available only in debug plugin builds.");
    }
    state.Require(SUCCEEDED(createHr) && created.fileSystem,
                  std::format(L"MTP journal orphan-temp cleanup: create selftest instance failed. hr=0x{:08X}", static_cast<unsigned long>(createHr)));
    if (FAILED(createHr) || ! created.fileSystem)
    {
        return false;
    }

    wil::com_ptr<IFileSystemIO> io;
    state.Require(CreateFileSystemIo(created.fileSystem, io), L"MTP journal orphan-temp cleanup: missing IFileSystemIO.");
    if (! io)
    {
        return false;
    }

    auto cleanup = wil::scope_exit([&]() noexcept
    {
        static_cast<void>(created.fileSystem->DeleteItem(finalPath.c_str(), FILESYSTEM_FLAG_NONE, nullptr, nullptr, nullptr));
        static_cast<void>(created.fileSystem->DeleteItem(tempPath.c_str(), FILESYSTEM_FLAG_NONE, nullptr, nullptr, nullptr));
        if (! journalPath.empty())
        {
            static_cast<void>(DeleteFileW(journalPath.c_str()));
        }
    });

    constexpr std::string_view kOriginalPayload = "original final object preserved during orphan temp cleanup";
    constexpr std::string_view kTempPayload     = "orphan temp payload from pre-delete crash window";
    state.Require(
        WritePluginFileText(io.get(), finalPath.c_str(), FILESYSTEM_FLAG_NONE, kOriginalPayload, state, L"MTP journal orphan-temp cleanup final write"),
        L"MTP journal orphan-temp cleanup: final write failed.");
    state.Require(WritePluginFileText(io.get(), tempPath.c_str(), FILESYSTEM_FLAG_NONE, kTempPayload, state, L"MTP journal orphan-temp cleanup temp write"),
                  L"MTP journal orphan-temp cleanup: temp write failed.");
    if (! state.failure.empty())
    {
        return false;
    }

    HRESULT hr = ensureDirectoryExists(journalDirectory);
    state.Require(SUCCEEDED(hr), std::format(L"MTP journal orphan-temp cleanup: ensure journal directory failed. hr=0x{:08X}", static_cast<unsigned long>(hr)));
    if (FAILED(hr))
    {
        return false;
    }

    const std::string journalJson = std::format(
        R"json({{"schemaVersion":1,"entries":[{{"phase":"planned","devicePuid":"{}","sourcePath":"","destinationPath":"{}","tempPath":"{}","declaredSizeBytes":{},"sourceTransmitHashHex":"manual-selftest","journalTimestampFileTimeUtc":1}}]}})json",
        narrowAscii(host),
        narrowAscii(finalPath),
        narrowAscii(tempPath),
        static_cast<unsigned long long>(kTempPayload.size()));
    hr = writeUtf8File(journalPath, journalJson);
    state.Require(SUCCEEDED(hr), std::format(L"MTP journal orphan-temp cleanup: write journal failed. hr=0x{:08X}", static_cast<unsigned long>(hr)));
    if (FAILED(hr))
    {
        return false;
    }

    hr = notifyInjectedJournal(host);
    state.Require(SUCCEEDED(hr),
                  std::format(L"MTP journal orphan-temp cleanup: cache invalidation failed. hr=0x{:08X}", static_cast<unsigned long>(hr)));
    if (FAILED(hr))
    {
        return false;
    }

    std::string readBack;
    state.Require(ReadPluginFileText(io.get(), finalPath.c_str(), readBack, state, L"MTP journal orphan-temp cleanup replay readback"),
                  L"MTP journal orphan-temp cleanup: replay readback failed.");
    state.Require(readBack == kOriginalPayload, L"MTP journal orphan-temp cleanup: replay replaced the final object instead of removing temp.");

    wil::com_ptr<IFileReader> tempReader;
    hr = io->CreateFileReader(tempPath.c_str(), tempReader.put());
    state.Require(hr == HRESULT_FROM_WIN32(ERROR_FILE_NOT_FOUND) && ! tempReader,
                  std::format(L"MTP journal orphan-temp cleanup: temp should be removed, got hr=0x{:08X}.", static_cast<unsigned long>(hr)));

    const DWORD journalAttributes = GetFileAttributesW(journalPath.c_str());
    state.Require(journalAttributes == INVALID_FILE_ATTRIBUTES && GetLastError() == ERROR_FILE_NOT_FOUND,
                  L"MTP journal orphan-temp cleanup: replay did not clear the host journal.");

    const auto entries      = SnapshotDirectoryEntries(created.fileSystem, baseFolder.c_str(), state, L"MTP journal orphan-temp cleanup camera");
    uint32_t finalNameCount = 0;
    for (const DirectoryEntrySnapshot& entry : entries)
    {
        if (entry.name == fileName)
        {
            ++finalNameCount;
        }
        state.Require(entry.name != tempName, std::format(L"MTP journal orphan-temp cleanup: leaked temp entry {} after replay.", entry.name));
    }
    state.Require(finalNameCount == 1u, std::format(L"MTP journal orphan-temp cleanup: expected one final entry after replay, got {}.", finalNameCount));

    return state.failure.empty();
});

SelfTest::RunCase(options,
                  suite,
                  L"mtp_overwrite_journal_replay_temp_cleanup_delete_failure_retries",
                  [&](SelfTest::CaseState& state) noexcept
{
    const std::wstring guid = MakeGuidText();
    state.Require(! guid.empty(), L"MTP journal temp cleanup retry: failed to generate unique file name.");
    if (guid.empty())
    {
        return false;
    }

    const std::wstring baseFolder   = L"/Fake Phone [devpuid:fake-device]/Internal Storage/DCIM/Camera";
    const std::wstring fileName     = L"journal-temp-cleanup-retry-" + guid + L".txt";
    const std::wstring tempName     = L".rs-mtp-overwrite-cleanup-retry-" + guid + L".tmp";
    const std::wstring finalPath    = baseFolder + L"/" + fileName;
    const std::wstring tempPath     = baseFolder + L"/" + tempName;
    const std::wstring host         = L"journal-temp-cleanup-retry-" + guid;
    const std::string fakeOptions   = std::format(R"json({{"deleteItemFailOncePath":"{}"}})json", narrowAscii(tempPath));
    const std::string configuration = std::format(R"json({{"readOnly":false,"host":"{}"}})json", narrowAscii(host));

    std::wstring previousLocalAppData;
    std::wstring localAppData;
    if (! AcquireMtpJournalLocalAppDataSandbox(
            state, L"mtp_journal_temp_retry", L"MTP journal temp cleanup retry", localAppData, previousLocalAppData))
    {
        return false;
    }
    const auto restoreLocalAppData = wil::scope_exit([&]() noexcept { RestoreMtpJournalLocalAppDataSandbox(previousLocalAppData); });

    const std::wstring journalDirectory = localAppData + L"\\RedSalamander\\PluginState\\FileSystemMtp\\" + std::format(L"{:016X}", stableDeviceHash(host));
    const std::wstring journalPath      = journalDirectory + L"\\overwrite-journal.json";

    CreatedFileSystemInstance created;
    const HRESULT createHr = TryCreateFakeMtpFileSystemInstance(fakeOptions, configuration, L"/", created);
    if (createHr == HRESULT_FROM_WIN32(ERROR_PROC_NOT_FOUND))
    {
        return state.Skip(L"MTP fake backend export is available only in debug plugin builds.");
    }
    state.Require(SUCCEEDED(createHr) && created.fileSystem,
                  std::format(L"MTP journal temp cleanup retry: create selftest instance failed. hr=0x{:08X}", static_cast<unsigned long>(createHr)));
    if (FAILED(createHr) || ! created.fileSystem)
    {
        return false;
    }

    wil::com_ptr<IFileSystemIO> io;
    state.Require(CreateFileSystemIo(created.fileSystem, io), L"MTP journal temp cleanup retry: missing IFileSystemIO.");
    if (! io)
    {
        return false;
    }

    auto cleanup = wil::scope_exit([&]() noexcept
    {
        static_cast<void>(created.fileSystem->DeleteItem(finalPath.c_str(), FILESYSTEM_FLAG_NONE, nullptr, nullptr, nullptr));
        static_cast<void>(created.fileSystem->DeleteItem(tempPath.c_str(), FILESYSTEM_FLAG_NONE, nullptr, nullptr, nullptr));
        if (! journalPath.empty())
        {
            static_cast<void>(DeleteFileW(journalPath.c_str()));
        }
    });

    constexpr std::string_view kOriginalPayload = "original final object preserved while temp cleanup retries";
    constexpr std::string_view kTempPayload     = "temp cleanup retry payload";
    state.Require(
        WritePluginFileText(io.get(), finalPath.c_str(), FILESYSTEM_FLAG_NONE, kOriginalPayload, state, L"MTP journal temp cleanup retry final write"),
        L"MTP journal temp cleanup retry: final write failed.");
    state.Require(WritePluginFileText(io.get(), tempPath.c_str(), FILESYSTEM_FLAG_NONE, kTempPayload, state, L"MTP journal temp cleanup retry temp write"),
                  L"MTP journal temp cleanup retry: temp write failed.");
    if (! state.failure.empty())
    {
        return false;
    }

    HRESULT hr = ensureDirectoryExists(journalDirectory);
    state.Require(SUCCEEDED(hr), std::format(L"MTP journal temp cleanup retry: ensure journal directory failed. hr=0x{:08X}", static_cast<unsigned long>(hr)));
    if (FAILED(hr))
    {
        return false;
    }

    const std::string journalJson = std::format(
        R"json({{"schemaVersion":1,"entries":[{{"phase":"planned","devicePuid":"{}","sourcePath":"","destinationPath":"{}","tempPath":"{}","declaredSizeBytes":{},"sourceTransmitHashHex":"manual-selftest","journalTimestampFileTimeUtc":1}}]}})json",
        narrowAscii(host),
        narrowAscii(finalPath),
        narrowAscii(tempPath),
        static_cast<unsigned long long>(kTempPayload.size()));
    hr = writeUtf8File(journalPath, journalJson);
    state.Require(SUCCEEDED(hr), std::format(L"MTP journal temp cleanup retry: write journal failed. hr=0x{:08X}", static_cast<unsigned long>(hr)));
    if (FAILED(hr))
    {
        return false;
    }

    hr = notifyInjectedJournal(host);
    state.Require(SUCCEEDED(hr),
                  std::format(L"MTP journal temp cleanup retry: cache invalidation failed. hr=0x{:08X}", static_cast<unsigned long>(hr)));
    if (FAILED(hr))
    {
        return false;
    }

    wil::com_ptr<IFileReader> failedReader;
    hr = io->CreateFileReader(finalPath.c_str(), failedReader.put());
    state.Require(
        hr == HRESULT_FROM_WIN32(ERROR_ACCESS_DENIED) && ! failedReader,
        std::format(L"MTP journal temp cleanup retry: first replay expected cleanup ERROR_ACCESS_DENIED, got hr=0x{:08X}.", static_cast<unsigned long>(hr)));

    DWORD journalAttributes = GetFileAttributesW(journalPath.c_str());
    state.Require(journalAttributes != INVALID_FILE_ATTRIBUTES && (journalAttributes & FILE_ATTRIBUTE_DIRECTORY) == 0,
                  L"MTP journal temp cleanup retry: journal was cleared after failed cleanup.");

    if (! state.failure.empty())
    {
        return false;
    }

    std::string finalReadBack;
    state.Require(ReadPluginFileText(io.get(), finalPath.c_str(), finalReadBack, state, L"MTP journal temp cleanup retry final retry"),
                  L"MTP journal temp cleanup retry: final retry readback failed.");
    state.Require(finalReadBack == kOriginalPayload, L"MTP journal temp cleanup retry: retry changed the final contents.");

    wil::com_ptr<IFileReader> tempReader;
    hr = io->CreateFileReader(tempPath.c_str(), tempReader.put());
    state.Require(hr == HRESULT_FROM_WIN32(ERROR_FILE_NOT_FOUND) && ! tempReader,
                  std::format(L"MTP journal temp cleanup retry: temp should be removed after retry, got hr=0x{:08X}.", static_cast<unsigned long>(hr)));

    journalAttributes = GetFileAttributesW(journalPath.c_str());
    state.Require(journalAttributes == INVALID_FILE_ATTRIBUTES && GetLastError() == ERROR_FILE_NOT_FOUND,
                  L"MTP journal temp cleanup retry: retry did not clear the host journal.");

    const auto entries      = SnapshotDirectoryEntries(created.fileSystem, baseFolder.c_str(), state, L"MTP journal temp cleanup retry camera");
    uint32_t finalNameCount = 0;
    for (const DirectoryEntrySnapshot& entry : entries)
    {
        if (entry.name == fileName)
        {
            ++finalNameCount;
        }
        state.Require(entry.name != tempName, std::format(L"MTP journal temp cleanup retry: leaked temp entry {} after retry.", entry.name));
    }
    state.Require(finalNameCount == 1u, std::format(L"MTP journal temp cleanup retry: expected one final entry after retry, got {}.", finalNameCount));

    return state.failure.empty();
});

SelfTest::RunCase(options,
                  suite,
                  L"mtp_overwrite_journal_recovers_committed_temp_without_tempPuid",
                  [&](SelfTest::CaseState& state) noexcept
{
    const std::wstring guid = MakeGuidText();
    state.Require(! guid.empty(), L"MTP journal no-tempPUID sweep: failed to generate unique file name.");
    if (guid.empty())
    {
        return false;
    }

    std::wstring previousLocalAppData;
    std::wstring localAppData;
    if (! AcquireMtpJournalLocalAppDataSandbox(
            state, L"mtp_journal_no_temp_puid", L"MTP journal no-tempPUID sweep", localAppData, previousLocalAppData))
    {
        return false;
    }
    const auto restoreLocalAppData = wil::scope_exit([&]() noexcept { RestoreMtpJournalLocalAppDataSandbox(previousLocalAppData); });

    const std::wstring baseFolder = L"/Fake Phone [devpuid:fake-device]/Internal Storage/DCIM/Camera";

    const auto runScenario = [&](bool ambiguous) noexcept -> bool
    {
        const std::wstring scenario         = ambiguous ? L"ambiguous" : L"single";
        const std::wstring host             = L"journal-no-temp-puid-" + scenario + L"-" + guid;
        const std::wstring journalDirectory = localAppData + L"\\RedSalamander\\PluginState\\FileSystemMtp\\" + std::format(L"{:016X}", stableDeviceHash(host));
        const std::wstring journalPath      = journalDirectory + L"\\overwrite-journal.json";
        const std::wstring finalName        = L"journal-no-temp-puid-" + scenario + L"-" + guid + L".txt";
        const std::wstring tempLeaf         = L"." + finalName + L".rs-mtp-overwrite-no-temp-puid-" + scenario + L"-" + guid + L".tmp";
        const std::wstring finalPath        = baseFolder + L"/" + finalName;
        const std::wstring journalTempPath  = baseFolder + L"/" + tempLeaf;
        const std::wstring candidateOneName = tempLeaf + L" [puid:manual-one]";
        const std::wstring candidateTwoName = tempLeaf + L" [puid:manual-two]";
        const std::wstring candidateOnePath = baseFolder + L"/" + candidateOneName;
        const std::wstring candidateTwoPath = baseFolder + L"/" + candidateTwoName;
        const std::string configuration     = std::format(R"json({{"readOnly":false,"host":"{}"}})json", narrowAscii(host));

        CreatedFileSystemInstance created;
        const HRESULT createHr = TryCreateFakeMtpFileSystemInstance("{}", configuration, L"/", created);
        if (createHr == HRESULT_FROM_WIN32(ERROR_PROC_NOT_FOUND))
        {
            return state.Skip(L"MTP fake backend export is available only in debug plugin builds.");
        }
        state.Require(
            SUCCEEDED(createHr) && created.fileSystem,
            std::format(L"MTP journal no-tempPUID sweep: create selftest instance for {} failed. hr=0x{:08X}", scenario, static_cast<unsigned long>(createHr)));
        if (FAILED(createHr) || ! created.fileSystem)
        {
            return false;
        }

        wil::com_ptr<IFileSystemIO> io;
        state.Require(CreateFileSystemIo(created.fileSystem, io), std::format(L"MTP journal no-tempPUID sweep: {} missing IFileSystemIO.", scenario));
        if (! io)
        {
            return false;
        }

        auto cleanup = wil::scope_exit([&]() noexcept
        {
            if (! journalPath.empty())
            {
                static_cast<void>(DeleteFileW(journalPath.c_str()));
            }
            static_cast<void>(created.fileSystem->DeleteItem(finalPath.c_str(), FILESYSTEM_FLAG_NONE, nullptr, nullptr, nullptr));
            static_cast<void>(created.fileSystem->DeleteItem(candidateOnePath.c_str(), FILESYSTEM_FLAG_NONE, nullptr, nullptr, nullptr));
            static_cast<void>(created.fileSystem->DeleteItem(candidateTwoPath.c_str(), FILESYSTEM_FLAG_NONE, nullptr, nullptr, nullptr));
        });

        constexpr std::string_view kOriginalPayload  = "original final object preserved during no-tempPUID sweep";
        constexpr std::string_view kCandidatePayload = "candidate temp object matched by journal metadata";
        state.Require(
            WritePluginFileText(io.get(), finalPath.c_str(), FILESYSTEM_FLAG_NONE, kOriginalPayload, state, L"MTP journal no-tempPUID sweep final write"),
            std::format(L"MTP journal no-tempPUID sweep: {} final write failed.", scenario));
        state.Require(
            WritePluginFileText(
                io.get(), candidateOnePath.c_str(), FILESYSTEM_FLAG_NONE, kCandidatePayload, state, L"MTP journal no-tempPUID sweep candidate one write"),
            std::format(L"MTP journal no-tempPUID sweep: {} candidate one write failed.", scenario));
        if (ambiguous)
        {
            state.Require(
                WritePluginFileText(
                    io.get(), candidateTwoPath.c_str(), FILESYSTEM_FLAG_NONE, kCandidatePayload, state, L"MTP journal no-tempPUID sweep candidate two write"),
                L"MTP journal no-tempPUID sweep: ambiguous candidate two write failed.");
        }
        if (! state.failure.empty())
        {
            return false;
        }

        HRESULT hr = ensureDirectoryExists(journalDirectory);
        state.Require(
            SUCCEEDED(hr),
            std::format(L"MTP journal no-tempPUID sweep: ensure journal directory for {} failed. hr=0x{:08X}", scenario, static_cast<unsigned long>(hr)));
        if (FAILED(hr))
        {
            return false;
        }

        const std::string journalJson = std::format(
            R"json({{"schemaVersion":1,"entries":[{{"phase":"committed","devicePuid":"{}","sourcePath":"","destinationPath":"{}","tempPath":"{}","declaredSizeBytes":{},"sourceTransmitHashHex":"manual-selftest","journalTimestampFileTimeUtc":1}}]}})json",
            narrowAscii(host),
            narrowAscii(finalPath),
            narrowAscii(journalTempPath),
            static_cast<unsigned long long>(kCandidatePayload.size()));
        hr = writeUtf8File(journalPath, journalJson);
        state.Require(SUCCEEDED(hr),
                      std::format(L"MTP journal no-tempPUID sweep: write journal for {} failed. hr=0x{:08X}", scenario, static_cast<unsigned long>(hr)));
        if (FAILED(hr))
        {
            return false;
        }

        hr = notifyInjectedJournal(host);
        state.Require(SUCCEEDED(hr),
                      std::format(L"MTP journal no-tempPUID sweep: cache invalidation for {} failed. hr=0x{:08X}",
                                  scenario,
                                  static_cast<unsigned long>(hr)));
        if (FAILED(hr))
        {
            return false;
        }

        std::string finalReadBack;
        state.Require(ReadPluginFileText(io.get(), finalPath.c_str(), finalReadBack, state, L"MTP journal no-tempPUID sweep final read"),
                      std::format(L"MTP journal no-tempPUID sweep: {} final read failed.", scenario));
        state.Require(finalReadBack == kOriginalPayload, std::format(L"MTP journal no-tempPUID sweep: {} changed final contents.", scenario));
        if (! state.failure.empty())
        {
            return false;
        }

        DWORD journalAttributes = GetFileAttributesW(journalPath.c_str());
        if (ambiguous)
        {
            state.Require(journalAttributes != INVALID_FILE_ATTRIBUTES && (journalAttributes & FILE_ATTRIBUTE_DIRECTORY) == 0,
                          L"MTP journal no-tempPUID sweep: ambiguous sweep did not retain the journal.");

            std::string candidateReadBack;
            state.Require(ReadPluginFileText(io.get(), candidateOnePath.c_str(), candidateReadBack, state, L"MTP journal no-tempPUID sweep candidate one read"),
                          L"MTP journal no-tempPUID sweep: ambiguous candidate one read failed.");
            state.Require(candidateReadBack == kCandidatePayload, L"MTP journal no-tempPUID sweep: ambiguous candidate one was changed.");
            candidateReadBack.clear();
            state.Require(ReadPluginFileText(io.get(), candidateTwoPath.c_str(), candidateReadBack, state, L"MTP journal no-tempPUID sweep candidate two read"),
                          L"MTP journal no-tempPUID sweep: ambiguous candidate two read failed.");
            state.Require(candidateReadBack == kCandidatePayload, L"MTP journal no-tempPUID sweep: ambiguous candidate two was changed.");
        }
        else
        {
            state.Require(journalAttributes == INVALID_FILE_ATTRIBUTES && GetLastError() == ERROR_FILE_NOT_FOUND,
                          L"MTP journal no-tempPUID sweep: exact sweep did not clear the host journal.");

            wil::com_ptr<IFileReader> candidateReader;
            hr = io->CreateFileReader(candidateOnePath.c_str(), candidateReader.put());
            state.Require(hr == HRESULT_FROM_WIN32(ERROR_FILE_NOT_FOUND) && ! candidateReader,
                          std::format(L"MTP journal no-tempPUID sweep: exact candidate should be deleted, got hr=0x{:08X}.", static_cast<unsigned long>(hr)));
        }

        const auto entries         = SnapshotDirectoryEntries(created.fileSystem, baseFolder.c_str(), state, L"MTP journal no-tempPUID sweep camera");
        uint32_t finalNameCount    = 0;
        uint32_t candidateOneCount = 0;
        uint32_t candidateTwoCount = 0;
        for (const DirectoryEntrySnapshot& entry : entries)
        {
            if (entry.name == finalName)
            {
                ++finalNameCount;
            }
            if (entry.name == candidateOneName)
            {
                ++candidateOneCount;
            }
            if (entry.name == candidateTwoName)
            {
                ++candidateTwoCount;
            }
        }
        state.Require(finalNameCount == 1u, std::format(L"MTP journal no-tempPUID sweep: {} expected one final entry, got {}.", scenario, finalNameCount));
        if (ambiguous)
        {
            state.Require(
                candidateOneCount == 1u && candidateTwoCount == 1u,
                std::format(L"MTP journal no-tempPUID sweep: ambiguous candidates changed. first={} second={}.", candidateOneCount, candidateTwoCount));
        }
        else
        {
            state.Require(candidateOneCount == 0u,
                          std::format(L"MTP journal no-tempPUID sweep: exact candidate remained visible count={}.", candidateOneCount));
        }

        return state.failure.empty();
    };

    if (! runScenario(false))
    {
        return false;
    }
    if (! runScenario(true))
    {
        return false;
    }

    return state.failure.empty();
});

SelfTest::RunCase(options,
                  suite,
                  L"mtp_overwrite_journal_replay_rename_rejection_is_bounded",
                  [&](SelfTest::CaseState& state) noexcept
{
    const std::wstring guid = MakeGuidText();
    state.Require(! guid.empty(), L"MTP journal rename rejection bound: failed to generate unique file name.");
    if (guid.empty())
    {
        return false;
    }

    const std::wstring baseFolder   = L"/Fake Phone [devpuid:fake-device]/Internal Storage/DCIM/Camera";
    const std::wstring finalName    = L"journal-rename-reject-" + guid + L".txt";
    const std::wstring tempName     = L".rs-mtp-overwrite-rename-reject-" + guid + L".tmp";
    const std::wstring finalPath    = baseFolder + L"/" + finalName;
    const std::wstring tempPath     = baseFolder + L"/" + tempName;
    const std::wstring host         = L"journal-rename-reject-" + guid;
    const std::string fakeOptions   = std::format(R"json({{"renameItemFailDestinationPath":"{}"}})json", narrowAscii(finalPath));
    const std::string configuration = std::format(R"json({{"readOnly":false,"host":"{}"}})json", narrowAscii(host));

    std::wstring previousLocalAppData;
    std::wstring localAppData;
    if (! AcquireMtpJournalLocalAppDataSandbox(
            state, L"mtp_journal_rename_reject", L"MTP journal rename rejection bound", localAppData, previousLocalAppData))
    {
        return false;
    }
    const auto restoreLocalAppData = wil::scope_exit([&]() noexcept { RestoreMtpJournalLocalAppDataSandbox(previousLocalAppData); });

    const std::wstring journalDirectory = localAppData + L"\\RedSalamander\\PluginState\\FileSystemMtp\\" + std::format(L"{:016X}", stableDeviceHash(host));
    const std::wstring journalPath      = journalDirectory + L"\\overwrite-journal.json";

    CreatedFileSystemInstance created;
    const HRESULT createHr = TryCreateFakeMtpFileSystemInstance(fakeOptions, configuration, L"/", created);
    if (createHr == HRESULT_FROM_WIN32(ERROR_PROC_NOT_FOUND))
    {
        return state.Skip(L"MTP fake backend export is available only in debug plugin builds.");
    }
    state.Require(SUCCEEDED(createHr) && created.fileSystem,
                  std::format(L"MTP journal rename rejection bound: create selftest instance failed. hr=0x{:08X}", static_cast<unsigned long>(createHr)));
    if (FAILED(createHr) || ! created.fileSystem)
    {
        return false;
    }

    wil::com_ptr<IFileSystemIO> io;
    state.Require(CreateFileSystemIo(created.fileSystem, io), L"MTP journal rename rejection bound: missing IFileSystemIO.");
    if (! io)
    {
        return false;
    }

    auto cleanup = wil::scope_exit([&]() noexcept
    {
        static_cast<void>(created.fileSystem->DeleteItem(finalPath.c_str(), FILESYSTEM_FLAG_NONE, nullptr, nullptr, nullptr));
        static_cast<void>(created.fileSystem->DeleteItem(tempPath.c_str(), FILESYSTEM_FLAG_NONE, nullptr, nullptr, nullptr));
        if (! journalPath.empty())
        {
            static_cast<void>(DeleteFileW(journalPath.c_str()));
        }
    });

    constexpr std::string_view kTempPayload = "verified temp payload left under temp name after bounded replay rejection";
    state.Require(WritePluginFileText(io.get(), tempPath.c_str(), FILESYSTEM_FLAG_NONE, kTempPayload, state, L"MTP journal rename rejection bound temp write"),
                  L"MTP journal rename rejection bound: temp write failed.");
    if (! state.failure.empty())
    {
        return false;
    }

    HRESULT hr = ensureDirectoryExists(journalDirectory);
    state.Require(SUCCEEDED(hr),
                  std::format(L"MTP journal rename rejection bound: ensure journal directory failed. hr=0x{:08X}", static_cast<unsigned long>(hr)));
    if (FAILED(hr))
    {
        return false;
    }

    const std::string journalJson = std::format(
        R"json({{"schemaVersion":1,"entries":[{{"phase":"planned","devicePuid":"{}","sourcePath":"","destinationPath":"{}","tempPath":"{}","declaredSizeBytes":{},"sourceTransmitHashHex":"manual-selftest","journalTimestampFileTimeUtc":1,"replayAttemptCount":0}}]}})json",
        narrowAscii(host),
        narrowAscii(finalPath),
        narrowAscii(tempPath),
        static_cast<unsigned long long>(kTempPayload.size()));
    hr = writeUtf8File(journalPath, journalJson);
    state.Require(SUCCEEDED(hr), std::format(L"MTP journal rename rejection bound: write journal failed. hr=0x{:08X}", static_cast<unsigned long>(hr)));
    if (FAILED(hr))
    {
        return false;
    }

    hr = notifyInjectedJournal(host);
    state.Require(SUCCEEDED(hr),
                  std::format(L"MTP journal rename rejection bound: cache invalidation failed. hr=0x{:08X}", static_cast<unsigned long>(hr)));
    if (FAILED(hr))
    {
        return false;
    }

    const auto requireJournalPresent = [&](std::wstring_view label) noexcept
    {
        const DWORD attributes = GetFileAttributesW(journalPath.c_str());
        state.Require(attributes != INVALID_FILE_ATTRIBUTES && (attributes & FILE_ATTRIBUTE_DIRECTORY) == 0,
                      std::format(L"MTP journal rename rejection bound: journal missing after {}.", label));
    };

    for (uint32_t attempt = 1; attempt < 3u; ++attempt)
    {
        wil::com_ptr<IFileReader> reader;
        hr = io->CreateFileReader(finalPath.c_str(), reader.put());
        state.Require(hr == HRESULT_FROM_WIN32(ERROR_ACCESS_DENIED) && ! reader,
                      std::format(L"MTP journal rename rejection bound: attempt {} expected ERROR_ACCESS_DENIED, got hr=0x{:08X}.",
                                  attempt,
                                  static_cast<unsigned long>(hr)));
        requireJournalPresent(std::format(L"retry attempt {}", attempt));
        if (! state.failure.empty())
        {
            return false;
        }
    }

    wil::com_ptr<IFileReader> reader;
    hr = io->CreateFileReader(finalPath.c_str(), reader.put());
    state.Require(hr == HRESULT_FROM_WIN32(ERROR_FILE_NOT_FOUND) && ! reader,
                  std::format(L"MTP journal rename rejection bound: terminal attempt should clear journal then miss final, got hr=0x{:08X}.",
                              static_cast<unsigned long>(hr)));

    DWORD journalAttributes = GetFileAttributesW(journalPath.c_str());
    state.Require(journalAttributes == INVALID_FILE_ATTRIBUTES && GetLastError() == ERROR_FILE_NOT_FOUND,
                  L"MTP journal rename rejection bound: terminal replay did not clear the host journal.");

    reader.reset();
    hr = io->CreateFileReader(finalPath.c_str(), reader.put());
    state.Require(
        hr == HRESULT_FROM_WIN32(ERROR_FILE_NOT_FOUND) && ! reader,
        std::format(L"MTP journal rename rejection bound: post-terminal read should not retry rename, got hr=0x{:08X}.", static_cast<unsigned long>(hr)));

    std::string tempReadBack;
    state.Require(ReadPluginFileText(io.get(), tempPath.c_str(), tempReadBack, state, L"MTP journal rename rejection bound temp readback"),
                  L"MTP journal rename rejection bound: temp readback failed after terminal replay.");
    state.Require(tempReadBack == kTempPayload, L"MTP journal rename rejection bound: terminal replay changed the retained temp payload.");

    const auto entries      = SnapshotDirectoryEntries(created.fileSystem, baseFolder.c_str(), state, L"MTP journal rename rejection bound camera");
    uint32_t finalNameCount = 0;
    uint32_t tempNameCount  = 0;
    for (const DirectoryEntrySnapshot& entry : entries)
    {
        if (entry.name == finalName)
        {
            ++finalNameCount;
        }
        if (entry.name == tempName)
        {
            ++tempNameCount;
        }
    }
    state.Require(finalNameCount == 0u,
                  std::format(L"MTP journal rename rejection bound: final entry should not exist after terminal replay, got {}.", finalNameCount));
    state.Require(tempNameCount == 1u, std::format(L"MTP journal rename rejection bound: expected one retained temp entry, got {}.", tempNameCount));

    return state.failure.empty();
});

SelfTest::RunCase(options,
                  suite,
                  L"mtp_overwrite_never_duplicates_or_halfwrites",
                  [&](SelfTest::CaseState& state) noexcept
{
    const std::wstring guid = MakeGuidText();
    state.Require(! guid.empty(), L"MTP overwrite safety matrix: failed to generate unique case name.");
    if (guid.empty())
    {
        return false;
    }

    std::wstring previousLocalAppData;
    std::wstring localAppData;
    if (! AcquireMtpJournalLocalAppDataSandbox(state, L"mtp_overwrite_safety_matrix", L"MTP overwrite safety matrix", localAppData, previousLocalAppData))
    {
        return false;
    }
    const auto restoreLocalAppData = wil::scope_exit([&]() noexcept { RestoreMtpJournalLocalAppDataSandbox(previousLocalAppData); });

    constexpr std::wstring_view kBaseFolder         = L"/Fake Phone [devpuid:fake-device]/Internal Storage/DCIM/Camera";
    constexpr std::string_view kOriginalPayload     = "original object that must not be half-written";
    constexpr std::string_view kReplacementPayload  = "replacement object that may become final only after safe recovery";
    constexpr std::string_view kTerminalTempPayload = "verified temp retained in terminal safe state";

    const auto commitOverwrite = [&](IFileSystemIO* io, const std::wstring& path, std::string_view payload, std::wstring_view label) noexcept -> HRESULT
    {
        wil::com_ptr<IFileWriter> writer;
        HRESULT hr = io->CreateFileWriter(path.c_str(), FILESYSTEM_FLAG_ALLOW_OVERWRITE, writer.put());
        state.Require(SUCCEEDED(hr) && writer,
                      std::format(L"MTP overwrite safety matrix: {} CreateFileWriter failed. hr=0x{:08X}", label, static_cast<unsigned long>(hr)));
        if (FAILED(hr) || ! writer)
        {
            return hr;
        }

        const unsigned long expectedBytes = static_cast<unsigned long>(payload.size());
        unsigned long written             = 0;
        hr                                = writer->Write(payload.data(), expectedBytes, &written);
        state.Require(SUCCEEDED(hr) && written == expectedBytes,
                      std::format(L"MTP overwrite safety matrix: {} Write failed. wrote={} expected={} hr=0x{:08X}",
                                  label,
                                  written,
                                  expectedBytes,
                                  static_cast<unsigned long>(hr)));
        if (FAILED(hr) || written != expectedBytes)
        {
            return hr;
        }

        return writer->Commit();
    };

    const auto requireWriteCalls = [&](IFileSystemIO* io, const std::wstring& path, uint64_t expected, std::wstring_view label) noexcept
    {
        const char* properties = nullptr;
        const HRESULT propsHr  = io->GetItemProperties(path.c_str(), &properties);
        state.Require(SUCCEEDED(propsHr) && properties != nullptr,
                      std::format(L"MTP overwrite safety matrix: {} GetItemProperties failed. hr=0x{:08X}", label, static_cast<unsigned long>(propsHr)));
        if (FAILED(propsHr) || properties == nullptr)
        {
            return;
        }

        const std::optional<uint64_t> parsed = extractJsonUInt(properties, "writeFileCalls");
        state.Require(parsed.has_value(), std::format(L"MTP overwrite safety matrix: {} missing writeFileCalls.", label));
        constexpr uint64_t kMissingCounter = (std::numeric_limits<uint64_t>::max)();
        state.Require(parsed.value_or(kMissingCounter) == expected,
                      std::format(L"MTP overwrite safety matrix: {} expected writeFileCalls={}, got {}.", label, expected, parsed.value_or(kMissingCounter)));
    };

    const auto requireReadBack = [&](IFileSystemIO* io, const std::wstring& path, std::string_view expected, std::wstring_view label) noexcept
    {
        std::string readBack;
        state.Require(ReadPluginFileText(io, path.c_str(), readBack, state, std::format(L"MTP overwrite safety matrix {} read", label)),
                      std::format(L"MTP overwrite safety matrix: {} readback failed.", label));
        state.Require(readBack == expected, std::format(L"MTP overwrite safety matrix: {} payload mismatch.", label));
    };

    const auto requireDirectoryState = [&](IFileSystem* fileSystem,
                                           std::wstring_view finalName,
                                           uint32_t expectedFinalCount,
                                           std::wstring_view tempName,
                                           uint32_t expectedTempCount,
                                           std::wstring_view label) noexcept
    {
        const auto entries  = SnapshotDirectoryEntries(fileSystem, kBaseFolder.data(), state, std::format(L"MTP overwrite safety matrix {}", label));
        uint32_t finalCount = 0;
        uint32_t tempCount  = 0;
        for (const DirectoryEntrySnapshot& entry : entries)
        {
            if (entry.name == finalName)
            {
                ++finalCount;
            }
            if (! tempName.empty())
            {
                if (entry.name == tempName)
                {
                    ++tempCount;
                }
            }
            else
            {
                state.Require(entry.name.find(L".rs-mtp-overwrite-") == std::wstring::npos,
                              std::format(L"MTP overwrite safety matrix: {} leaked temp entry {}.", label, entry.name));
            }
        }

        state.Require(finalCount == expectedFinalCount,
                      std::format(L"MTP overwrite safety matrix: {} expected {} final entries, got {}.", label, expectedFinalCount, finalCount));
        if (! tempName.empty())
        {
            state.Require(tempCount == expectedTempCount,
                          std::format(L"MTP overwrite safety matrix: {} expected {} temp entries, got {}.", label, expectedTempCount, tempCount));
        }
    };

    bool fakeBackendUnavailable = false;
    const auto createInstance =
        [&](std::string_view fakeOptions, std::wstring_view host, CreatedFileSystemInstance& created, wil::com_ptr<IFileSystemIO>& io) noexcept -> bool
    {
        const std::string configuration = std::format(R"json({{"readOnly":false,"host":"{}"}})json", narrowAscii(host));
        const HRESULT createHr          = TryCreateFakeMtpFileSystemInstance(fakeOptions, configuration, L"/", created);
        if (createHr == HRESULT_FROM_WIN32(ERROR_PROC_NOT_FOUND))
        {
            fakeBackendUnavailable = true;
            state.Skip(L"MTP fake backend export is available only in debug plugin builds.");
            return false;
        }
        state.Require(SUCCEEDED(createHr) && created.fileSystem,
                      std::format(L"MTP overwrite safety matrix: create instance for {} failed. hr=0x{:08X}", host, static_cast<unsigned long>(createHr)));
        if (FAILED(createHr) || ! created.fileSystem)
        {
            return false;
        }

        state.Require(CreateFileSystemIo(created.fileSystem, io), std::format(L"MTP overwrite safety matrix: {} missing IFileSystemIO.", host));
        return io != nullptr && state.failure.empty();
    };

    const auto journalPathForHost = [&](std::wstring_view host, std::wstring& journalDirectory, std::wstring& journalPath) noexcept -> bool
    {
        journalDirectory = localAppData + L"\\RedSalamander\\PluginState\\FileSystemMtp\\" + std::format(L"{:016X}", stableDeviceHash(host));
        journalPath      = journalDirectory + L"\\overwrite-journal.json";
        const HRESULT hr = ensureDirectoryExists(journalDirectory);
        state.Require(SUCCEEDED(hr),
                      std::format(L"MTP overwrite safety matrix: ensure journal directory for {} failed. hr=0x{:08X}", host, static_cast<unsigned long>(hr)));
        return SUCCEEDED(hr);
    };

    const auto writeJournal = [&](std::wstring_view host,
                                  const std::wstring& finalPath,
                                  const std::wstring& tempPath,
                                  uint64_t declaredSize,
                                  uint32_t replayAttemptCount,
                                  std::wstring& journalPath) noexcept -> bool
    {
        std::wstring journalDirectory;
        if (! journalPathForHost(host, journalDirectory, journalPath))
        {
            return false;
        }

        const std::string journalJson = std::format(
            R"json({{"schemaVersion":1,"entries":[{{"phase":"planned","devicePuid":"{}","sourcePath":"","destinationPath":"{}","tempPath":"{}","declaredSizeBytes":{},"sourceTransmitHashHex":"manual-selftest","journalTimestampFileTimeUtc":1,"replayAttemptCount":{}}}]}})json",
            narrowAscii(host),
            narrowAscii(finalPath),
            narrowAscii(tempPath),
            declaredSize,
            replayAttemptCount);
        const HRESULT hr = writeUtf8File(journalPath, journalJson);
        state.Require(SUCCEEDED(hr),
                      std::format(L"MTP overwrite safety matrix: write journal for {} failed. hr=0x{:08X}", host, static_cast<unsigned long>(hr)));
        if (FAILED(hr))
        {
            return false;
        }

        const HRESULT invalidateHr = notifyInjectedJournal(host);
        state.Require(SUCCEEDED(invalidateHr),
                      std::format(L"MTP overwrite safety matrix: cache invalidation for {} failed. hr=0x{:08X}",
                                  host,
                                  static_cast<unsigned long>(invalidateHr)));
        return SUCCEEDED(invalidateHr);
    };

    {
        const std::wstring host = L"safety-commit-failure-" + guid;
        CreatedFileSystemInstance created;
        wil::com_ptr<IFileSystemIO> io;
        if (! createInstance(R"json({"writeFileFailOncePathContains":".rs-mtp-overwrite-"})json", host, created, io))
        {
            return fakeBackendUnavailable;
        }

        const std::wstring fileName = L"safety-commit-failure-" + guid + L".txt";
        const std::wstring path     = std::wstring(kBaseFolder) + L"/" + fileName;
        auto cleanup                = wil::scope_exit([&]() noexcept
        { static_cast<void>(created.fileSystem->DeleteItem(path.c_str(), FILESYSTEM_FLAG_NONE, nullptr, nullptr, nullptr)); });

        state.Require(
            WritePluginFileText(io.get(), path.c_str(), FILESYSTEM_FLAG_NONE, kOriginalPayload, state, L"MTP overwrite safety matrix commit-failure initial"),
            L"MTP overwrite safety matrix: commit-failure initial write failed.");
        const HRESULT hr = commitOverwrite(io.get(), path, kReplacementPayload, L"commit failure");
        state.Require(hr == HRESULT_FROM_WIN32(ERROR_WRITE_FAULT),
                      std::format(L"MTP overwrite safety matrix: commit failure expected ERROR_WRITE_FAULT, got hr=0x{:08X}.", static_cast<unsigned long>(hr)));
        requireReadBack(io.get(), path, kOriginalPayload, L"commit failure");
        requireDirectoryState(created.fileSystem.get(), fileName, 1u, {}, 0u, L"commit failure");
        if (! state.failure.empty())
        {
            return false;
        }
    }

    {
        const std::wstring host = L"safety-journal-write-failure-" + guid;
        CreatedFileSystemInstance created;
        wil::com_ptr<IFileSystemIO> io;
        if (! createInstance("{}", host, created, io))
        {
            return fakeBackendUnavailable;
        }

        wil::com_ptr<IInformations> informations;
        const HRESULT qiHr = created.fileSystem->QueryInterface(__uuidof(IInformations), informations.put_void());
        state.Require(SUCCEEDED(qiHr) && informations,
                      std::format(L"MTP overwrite safety matrix: journal-write failure missing IInformations. hr=0x{:08X}", static_cast<unsigned long>(qiHr)));
        if (! informations)
        {
            return false;
        }
        const std::string configuration = std::format(R"json({{"readOnly":false,"host":"{}","failOverwriteJournalWrites":true}})json", narrowAscii(host));
        const HRESULT configHr          = informations->SetConfiguration(configuration.c_str());
        state.Require(
            SUCCEEDED(configHr),
            std::format(L"MTP overwrite safety matrix: journal-write failure SetConfiguration failed. hr=0x{:08X}", static_cast<unsigned long>(configHr)));
        if (FAILED(configHr))
        {
            return false;
        }

        const std::wstring fileName = L"safety-journal-write-failure-" + guid + L".txt";
        const std::wstring path     = std::wstring(kBaseFolder) + L"/" + fileName;
        auto cleanup                = wil::scope_exit([&]() noexcept
        { static_cast<void>(created.fileSystem->DeleteItem(path.c_str(), FILESYSTEM_FLAG_NONE, nullptr, nullptr, nullptr)); });

        state.Require(
            WritePluginFileText(io.get(), path.c_str(), FILESYSTEM_FLAG_NONE, kOriginalPayload, state, L"MTP overwrite safety matrix journal-write initial"),
            L"MTP overwrite safety matrix: journal-write failure initial write failed.");
        requireWriteCalls(io.get(), path, 1u, L"journal-write failure before overwrite");
        const HRESULT hr = commitOverwrite(io.get(), path, kReplacementPayload, L"journal write failure");
        state.Require(
            hr == HRESULT_FROM_WIN32(ERROR_ACCESS_DENIED),
            std::format(L"MTP overwrite safety matrix: journal write failure expected ERROR_ACCESS_DENIED, got hr=0x{:08X}.", static_cast<unsigned long>(hr)));
        requireWriteCalls(io.get(), path, 1u, L"journal-write failure after overwrite");
        requireReadBack(io.get(), path, kOriginalPayload, L"journal-write failure");
        requireDirectoryState(created.fileSystem.get(), fileName, 1u, {}, 0u, L"journal-write failure");
        if (! state.failure.empty())
        {
            return false;
        }
    }

    {
        const std::wstring host = L"safety-empty-temp-puid-" + guid;
        CreatedFileSystemInstance created;
        wil::com_ptr<IFileSystemIO> io;
        if (! createInstance(R"json({"omitPersistentIdForCreatedFiles":true})json", host, created, io))
        {
            return fakeBackendUnavailable;
        }

        const std::wstring fileName = L"safety-empty-temp-puid-" + guid + L".txt";
        const std::wstring path     = std::wstring(kBaseFolder) + L"/" + fileName;
        auto cleanup                = wil::scope_exit([&]() noexcept
        { static_cast<void>(created.fileSystem->DeleteItem(path.c_str(), FILESYSTEM_FLAG_NONE, nullptr, nullptr, nullptr)); });

        state.Require(
            WritePluginFileText(io.get(), path.c_str(), FILESYSTEM_FLAG_NONE, kOriginalPayload, state, L"MTP overwrite safety matrix empty-puid initial"),
            L"MTP overwrite safety matrix: empty-tempPUID initial write failed.");
        const HRESULT firstHr = commitOverwrite(io.get(), path, kReplacementPayload, L"empty tempPUID first");
        state.Require(
            firstHr == HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED),
            std::format(L"MTP overwrite safety matrix: empty tempPUID expected ERROR_NOT_SUPPORTED, got hr=0x{:08X}.", static_cast<unsigned long>(firstHr)));
        requireReadBack(io.get(), path, kOriginalPayload, L"empty tempPUID first");
        requireDirectoryState(created.fileSystem.get(), fileName, 1u, {}, 0u, L"empty tempPUID first");
        requireWriteCalls(io.get(), path, 2u, L"empty tempPUID after first failure");

        const HRESULT secondHr = commitOverwrite(io.get(), path, kReplacementPayload, L"empty tempPUID second");
        state.Require(secondHr == HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED),
                      std::format(L"MTP overwrite safety matrix: empty tempPUID second expected ERROR_NOT_SUPPORTED, got hr=0x{:08X}.",
                                  static_cast<unsigned long>(secondHr)));
        requireWriteCalls(io.get(), path, 2u, L"empty tempPUID after blocked retry");
        requireReadBack(io.get(), path, kOriginalPayload, L"empty tempPUID second");
        requireDirectoryState(created.fileSystem.get(), fileName, 1u, {}, 0u, L"empty tempPUID second");
        if (! state.failure.empty())
        {
            return false;
        }
    }

    {
        const std::wstring fileName   = L"safety-rename-recovery-" + guid + L".txt";
        const std::wstring path       = std::wstring(kBaseFolder) + L"/" + fileName;
        const std::wstring host       = L"safety-rename-recovery-" + guid;
        const std::string fakeOptions = std::format(R"json({{"renameItemFailOnceDestinationPath":"{}"}})json", narrowAscii(path));
        CreatedFileSystemInstance created;
        wil::com_ptr<IFileSystemIO> io;
        if (! createInstance(fakeOptions, host, created, io))
        {
            return fakeBackendUnavailable;
        }

        auto cleanup = wil::scope_exit([&]() noexcept
        { static_cast<void>(created.fileSystem->DeleteItem(path.c_str(), FILESYSTEM_FLAG_NONE, nullptr, nullptr, nullptr)); });

        state.Require(WritePluginFileText(io.get(), path.c_str(), FILESYSTEM_FLAG_NONE, kOriginalPayload, state, L"MTP overwrite safety matrix rename initial"),
                      L"MTP overwrite safety matrix: rename recovery initial write failed.");
        const HRESULT failedHr = commitOverwrite(io.get(), path, kReplacementPayload, L"rename recovery");
        state.Require(
            failedHr == HRESULT_FROM_WIN32(ERROR_ACCESS_DENIED),
            std::format(L"MTP overwrite safety matrix: rename recovery expected ERROR_ACCESS_DENIED, got hr=0x{:08X}.", static_cast<unsigned long>(failedHr)));
        requireReadBack(io.get(), path, kReplacementPayload, L"rename recovery replay");
        requireDirectoryState(created.fileSystem.get(), fileName, 1u, {}, 0u, L"rename recovery replay");
        if (! state.failure.empty())
        {
            return false;
        }
    }

    {
        const std::wstring host = L"safety-after-commit-orphan-" + guid;
        CreatedFileSystemInstance created;
        wil::com_ptr<IFileSystemIO> io;
        if (! createInstance("{}", host, created, io))
        {
            return fakeBackendUnavailable;
        }

        const std::wstring finalName = L"safety-after-commit-orphan-" + guid + L".txt";
        const std::wstring tempName  = L".rs-mtp-overwrite-safety-orphan-" + guid + L".tmp";
        const std::wstring finalPath = std::wstring(kBaseFolder) + L"/" + finalName;
        const std::wstring tempPath  = std::wstring(kBaseFolder) + L"/" + tempName;
        std::wstring journalPath;
        auto cleanup = wil::scope_exit([&]() noexcept
        {
            if (! journalPath.empty())
            {
                static_cast<void>(DeleteFileW(journalPath.c_str()));
            }
            static_cast<void>(created.fileSystem->DeleteItem(finalPath.c_str(), FILESYSTEM_FLAG_NONE, nullptr, nullptr, nullptr));
            static_cast<void>(created.fileSystem->DeleteItem(tempPath.c_str(), FILESYSTEM_FLAG_NONE, nullptr, nullptr, nullptr));
        });

        state.Require(
            WritePluginFileText(io.get(), finalPath.c_str(), FILESYSTEM_FLAG_NONE, kOriginalPayload, state, L"MTP overwrite safety matrix orphan final"),
            L"MTP overwrite safety matrix: orphan final write failed.");
        state.Require(
            WritePluginFileText(io.get(), tempPath.c_str(), FILESYSTEM_FLAG_NONE, kReplacementPayload, state, L"MTP overwrite safety matrix orphan temp"),
            L"MTP overwrite safety matrix: orphan temp write failed.");
        if (! writeJournal(host, finalPath, tempPath, static_cast<uint64_t>(kReplacementPayload.size()), 0u, journalPath))
        {
            return false;
        }

        requireReadBack(io.get(), finalPath, kOriginalPayload, L"after-commit orphan replay");
        wil::com_ptr<IFileReader> tempReader;
        const HRESULT tempHr = io->CreateFileReader(tempPath.c_str(), tempReader.put());
        state.Require(tempHr == HRESULT_FROM_WIN32(ERROR_FILE_NOT_FOUND) && ! tempReader,
                      std::format(L"MTP overwrite safety matrix: orphan temp should be removed, got hr=0x{:08X}.", static_cast<unsigned long>(tempHr)));
        requireDirectoryState(created.fileSystem.get(), finalName, 1u, tempName, 0u, L"after-commit orphan replay");
        if (! state.failure.empty())
        {
            return false;
        }
    }

    {
        const std::wstring host       = L"safety-rename-terminal-" + guid;
        const std::wstring finalName  = L"safety-rename-terminal-" + guid + L".txt";
        const std::wstring tempName   = L".rs-mtp-overwrite-safety-terminal-" + guid + L".tmp";
        const std::wstring finalPath  = std::wstring(kBaseFolder) + L"/" + finalName;
        const std::wstring tempPath   = std::wstring(kBaseFolder) + L"/" + tempName;
        const std::string fakeOptions = std::format(R"json({{"renameItemFailDestinationPath":"{}"}})json", narrowAscii(finalPath));
        CreatedFileSystemInstance created;
        wil::com_ptr<IFileSystemIO> io;
        if (! createInstance(fakeOptions, host, created, io))
        {
            return fakeBackendUnavailable;
        }

        std::wstring journalPath;
        auto cleanup = wil::scope_exit([&]() noexcept
        {
            if (! journalPath.empty())
            {
                static_cast<void>(DeleteFileW(journalPath.c_str()));
            }
            static_cast<void>(created.fileSystem->DeleteItem(finalPath.c_str(), FILESYSTEM_FLAG_NONE, nullptr, nullptr, nullptr));
            static_cast<void>(created.fileSystem->DeleteItem(tempPath.c_str(), FILESYSTEM_FLAG_NONE, nullptr, nullptr, nullptr));
        });

        state.Require(
            WritePluginFileText(io.get(), tempPath.c_str(), FILESYSTEM_FLAG_NONE, kTerminalTempPayload, state, L"MTP overwrite safety matrix terminal temp"),
            L"MTP overwrite safety matrix: terminal temp write failed.");
        if (! writeJournal(host, finalPath, tempPath, static_cast<uint64_t>(kTerminalTempPayload.size()), 0u, journalPath))
        {
            return false;
        }

        for (uint32_t attempt = 1; attempt < 3u; ++attempt)
        {
            wil::com_ptr<IFileReader> reader;
            const HRESULT hr = io->CreateFileReader(finalPath.c_str(), reader.put());
            state.Require(hr == HRESULT_FROM_WIN32(ERROR_ACCESS_DENIED) && ! reader,
                          std::format(L"MTP overwrite safety matrix: terminal retry {} expected ERROR_ACCESS_DENIED, got hr=0x{:08X}.",
                                      attempt,
                                      static_cast<unsigned long>(hr)));
        }

        wil::com_ptr<IFileReader> reader;
        const HRESULT terminalHr = io->CreateFileReader(finalPath.c_str(), reader.put());
        state.Require(
            terminalHr == HRESULT_FROM_WIN32(ERROR_FILE_NOT_FOUND) && ! reader,
            std::format(L"MTP overwrite safety matrix: terminal retry expected final missing, got hr=0x{:08X}.", static_cast<unsigned long>(terminalHr)));
        requireReadBack(io.get(), tempPath, kTerminalTempPayload, L"rename terminal temp");
        requireDirectoryState(created.fileSystem.get(), finalName, 0u, tempName, 1u, L"rename terminal");
        if (! state.failure.empty())
        {
            return false;
        }
    }

    return state.failure.empty();
});

SelfTest::RunCase(options,
                  suite,
                  L"mtp_overwrite_verify_input_by_source_kind",
                  [&](SelfTest::CaseState& state) noexcept
{
    const auto widenAscii = [](std::string_view text) { return std::wstring(text.begin(), text.end()); };

    struct VerifyInputExpectation
    {
        std::string_view level;
        bool move                      = false;
        uint32_t expectedReadFileCalls = 0;
        uint32_t expectedCopyItemCalls = 0;
    };

    constexpr std::array<VerifyInputExpectation, 4> kExpectations = {{
        {.level = "transmitHash", .move = false, .expectedReadFileCalls = 0, .expectedCopyItemCalls = 1},
        {.level = "transmitHash", .move = true, .expectedReadFileCalls = 0, .expectedCopyItemCalls = 1},
        {.level = "deviceReread", .move = false, .expectedReadFileCalls = 2, .expectedCopyItemCalls = 1},
        {.level = "deviceReread", .move = true, .expectedReadFileCalls = 2, .expectedCopyItemCalls = 1},
    }};

    const std::wstring guid = MakeGuidText();
    state.Require(! guid.empty(), L"MTP overwrite source-kind verify: failed to generate unique case name.");
    if (guid.empty())
    {
        return false;
    }

    constexpr std::string_view kSourcePayload = "device source overwrite verification payload";
    constexpr std::string_view kOldPayload    = "old destination payload";

    for (const VerifyInputExpectation& expectation : kExpectations)
    {
        CreatedFileSystemInstance created;
        const std::string configuration = std::format(R"json({{"readOnly":false,"byteVerifyOnOverwrite":"{}"}})json", expectation.level);
        const HRESULT createHr          = TryCreateFakeMtpFileSystemInstance("{}", configuration, L"/", created);
        if (createHr == HRESULT_FROM_WIN32(ERROR_PROC_NOT_FOUND))
        {
            return state.Skip(L"MTP fake backend export is available only in debug plugin builds.");
        }
        state.Require(SUCCEEDED(createHr) && created.fileSystem,
                      std::format(L"MTP overwrite source-kind verify: create selftest instance for {} failed. hr=0x{:08X}",
                                  widenAscii(expectation.level),
                                  static_cast<unsigned long>(createHr)));
        if (FAILED(createHr) || ! created.fileSystem)
        {
            return false;
        }

        wil::com_ptr<IFileSystemIO> io;
        state.Require(CreateFileSystemIo(created.fileSystem, io),
                      std::format(L"MTP overwrite source-kind verify: {} missing IFileSystemIO.", widenAscii(expectation.level)));
        if (! io)
        {
            return false;
        }

        const std::wstring levelName(expectation.level.begin(), expectation.level.end());
        const std::wstring opName = expectation.move ? L"move" : L"copy";
        const std::wstring sourcePath =
            L"/Fake Phone [devpuid:fake-device]/Internal Storage/DCIM/Camera/source-kind-" + opName + L"-" + levelName + L"-" + guid + L"-source.txt";
        const std::wstring destinationPath =
            L"/Fake Phone [devpuid:fake-device]/Internal Storage/DCIM/Camera/source-kind-" + opName + L"-" + levelName + L"-" + guid + L"-dest.txt";
        auto cleanup = wil::scope_exit([&]() noexcept
        {
            static_cast<void>(created.fileSystem->DeleteItem(sourcePath.c_str(), FILESYSTEM_FLAG_NONE, nullptr, nullptr, nullptr));
            static_cast<void>(created.fileSystem->DeleteItem(destinationPath.c_str(), FILESYSTEM_FLAG_NONE, nullptr, nullptr, nullptr));
        });

        state.Require(WritePluginFileText(io.get(), sourcePath.c_str(), FILESYSTEM_FLAG_NONE, kSourcePayload, state, L"MTP overwrite source-kind source write"),
                      std::format(L"MTP overwrite source-kind verify: {} {} source write failed.", widenAscii(expectation.level), opName));
        state.Require(
            WritePluginFileText(io.get(), destinationPath.c_str(), FILESYSTEM_FLAG_NONE, kOldPayload, state, L"MTP overwrite source-kind destination write"),
            std::format(L"MTP overwrite source-kind verify: {} {} destination write failed.", widenAscii(expectation.level), opName));
        if (! state.failure.empty())
        {
            return false;
        }

        const HRESULT opHr =
            expectation.move
                ? created.fileSystem->MoveItem(sourcePath.c_str(), destinationPath.c_str(), FILESYSTEM_FLAG_ALLOW_OVERWRITE, nullptr, nullptr, nullptr)
                : created.fileSystem->CopyItem(sourcePath.c_str(), destinationPath.c_str(), FILESYSTEM_FLAG_ALLOW_OVERWRITE, nullptr, nullptr, nullptr);
        state.Require(SUCCEEDED(opHr),
                      std::format(L"MTP overwrite source-kind verify: {} {} overwrite failed. hr=0x{:08X}",
                                  widenAscii(expectation.level),
                                  opName,
                                  static_cast<unsigned long>(opHr)));
        if (FAILED(opHr))
        {
            return false;
        }

        const char* properties = nullptr;
        const HRESULT propsHr  = io->GetItemProperties(destinationPath.c_str(), &properties);
        state.Require(SUCCEEDED(propsHr) && properties != nullptr,
                      std::format(L"MTP overwrite source-kind verify: {} {} GetItemProperties failed. hr=0x{:08X}",
                                  widenAscii(expectation.level),
                                  opName,
                                  static_cast<unsigned long>(propsHr)));
        if (SUCCEEDED(propsHr) && properties)
        {
            const std::string_view props(properties);
            const std::string expectedReads  = std::format(R"json("readFileCalls":{})json", expectation.expectedReadFileCalls);
            const std::string expectedSizes  = R"json("fileSizeCalls":0)json";
            const std::string expectedCopies = std::format(R"json("copyItemCalls":{})json", expectation.expectedCopyItemCalls);
            const std::string expectedLastRead =
                std::format(R"json("lastReadBytes":{})json", expectation.expectedReadFileCalls == 0 ? 0u : static_cast<uint32_t>(kSourcePayload.size()));

            state.Require(
                props.find(expectedReads) != std::string_view::npos,
                std::format(L"MTP overwrite source-kind verify: {} {} expected token {}.", widenAscii(expectation.level), opName, widenAscii(expectedReads)));
            state.Require(
                props.find(expectedSizes) != std::string_view::npos,
                std::format(L"MTP overwrite source-kind verify: {} {} should not use staged-writer size verification.", widenAscii(expectation.level), opName));
            state.Require(
                props.find(expectedCopies) != std::string_view::npos,
                std::format(L"MTP overwrite source-kind verify: {} {} expected token {}.", widenAscii(expectation.level), opName, widenAscii(expectedCopies)));
            state.Require(
                props.find(expectedLastRead) != std::string_view::npos,
                std::format(
                    L"MTP overwrite source-kind verify: {} {} expected token {}.", widenAscii(expectation.level), opName, widenAscii(expectedLastRead)));
        }

        std::string readBack;
        state.Require(ReadPluginFileText(io.get(), destinationPath.c_str(), readBack, state, L"MTP overwrite source-kind verify readback"),
                      std::format(L"MTP overwrite source-kind verify: {} {} readback failed.", widenAscii(expectation.level), opName));
        state.Require(readBack == kSourcePayload,
                      std::format(L"MTP overwrite source-kind verify: {} {} destination contents did not match.", widenAscii(expectation.level), opName));
    }

    return state.failure.empty();
});

SelfTest::RunCase(options,
                  suite,
                  L"mtp_copy_move_overwrite_uses_temp_puid_swap",
                  [&](SelfTest::CaseState& state) noexcept
{
    const auto extractJsonString = [](std::string_view json, std::string_view key) noexcept -> std::string
    {
        const std::string token = std::format(R"json("{}":")json", key);
        const size_t start      = json.find(token);
        if (start == std::string_view::npos)
        {
            return {};
        }

        const size_t valueStart = start + token.size();
        const size_t valueEnd   = json.find('"', valueStart);
        if (valueEnd == std::string_view::npos)
        {
            return {};
        }

        return std::string(json.substr(valueStart, valueEnd - valueStart));
    };
    const auto widenAscii = [](std::string_view text) { return std::wstring(text.begin(), text.end()); };

    const std::wstring guid = MakeGuidText();
    state.Require(! guid.empty(), L"MTP copy/move overwrite swap: failed to generate unique file name.");
    if (guid.empty())
    {
        return false;
    }

    const std::wstring baseFolder             = L"/Fake Phone [devpuid:fake-device]/Internal Storage/DCIM/Camera";
    constexpr std::array<bool, 2> kMoveValues = {{false, true}};
    for (const bool move : kMoveValues)
    {
        CreatedFileSystemInstance created;
        const HRESULT createHr = TryCreateFakeMtpFileSystemInstance("{}", R"json({"readOnly":false})json", L"/", created);
        if (createHr == HRESULT_FROM_WIN32(ERROR_PROC_NOT_FOUND))
        {
            return state.Skip(L"MTP fake backend export is available only in debug plugin builds.");
        }
        state.Require(SUCCEEDED(createHr) && created.fileSystem,
                      std::format(L"MTP copy/move overwrite swap: create selftest instance failed. hr=0x{:08X}", static_cast<unsigned long>(createHr)));
        if (FAILED(createHr) || ! created.fileSystem)
        {
            return false;
        }

        wil::com_ptr<IFileSystemIO> io;
        state.Require(CreateFileSystemIo(created.fileSystem, io), L"MTP copy/move overwrite swap: missing IFileSystemIO.");
        if (! io)
        {
            return false;
        }

        const std::wstring opName          = move ? L"move" : L"copy";
        const std::wstring sourcePath      = baseFolder + L"/copy-move-swap-" + opName + L"-" + guid + L"-source.txt";
        const std::wstring destName        = L"copy-move-swap-" + opName + L"-" + guid + L"-dest.txt";
        const std::wstring destinationPath = baseFolder + L"/" + destName;
        auto cleanup                       = wil::scope_exit([&]() noexcept
        {
            static_cast<void>(created.fileSystem->DeleteItem(sourcePath.c_str(), FILESYSTEM_FLAG_NONE, nullptr, nullptr, nullptr));
            static_cast<void>(created.fileSystem->DeleteItem(destinationPath.c_str(), FILESYSTEM_FLAG_NONE, nullptr, nullptr, nullptr));
        });

        constexpr std::string_view kSourcePayload = "copy/move source payload that should become destination";
        constexpr std::string_view kOldPayload    = "copy/move destination payload to be replaced";
        state.Require(
            WritePluginFileText(io.get(), sourcePath.c_str(), FILESYSTEM_FLAG_NONE, kSourcePayload, state, L"MTP copy/move overwrite swap source write"),
            std::format(L"MTP copy/move overwrite swap: {} source write failed.", opName));
        state.Require(
            WritePluginFileText(io.get(), destinationPath.c_str(), FILESYSTEM_FLAG_NONE, kOldPayload, state, L"MTP copy/move overwrite swap dest write"),
            std::format(L"MTP copy/move overwrite swap: {} destination write failed.", opName));
        if (! state.failure.empty())
        {
            return false;
        }

        const auto readPersistentId = [&](const std::wstring& path, std::string& persistentId, std::wstring_view label) noexcept
        {
            persistentId.clear();
            const char* properties = nullptr;
            const HRESULT propsHr  = io->GetItemProperties(path.c_str(), &properties);
            state.Require(
                SUCCEEDED(propsHr) && properties != nullptr,
                std::format(L"MTP copy/move overwrite swap: {} {} GetItemProperties failed. hr=0x{:08X}", opName, label, static_cast<unsigned long>(propsHr)));
            if (FAILED(propsHr) || properties == nullptr)
            {
                return;
            }

            persistentId = extractJsonString(properties, "persistentId");
            state.Require(! persistentId.empty(), std::format(L"MTP copy/move overwrite swap: {} {} did not expose a persistent id.", opName, label));
        };

        std::string oldDestinationPuid;
        readPersistentId(destinationPath, oldDestinationPuid, L"old destination");
        if (! state.failure.empty())
        {
            return false;
        }

        const HRESULT opHr =
            move ? created.fileSystem->MoveItem(sourcePath.c_str(), destinationPath.c_str(), FILESYSTEM_FLAG_ALLOW_OVERWRITE, nullptr, nullptr, nullptr)
                 : created.fileSystem->CopyItem(sourcePath.c_str(), destinationPath.c_str(), FILESYSTEM_FLAG_ALLOW_OVERWRITE, nullptr, nullptr, nullptr);
        state.Require(SUCCEEDED(opHr),
                      std::format(L"MTP copy/move overwrite swap: {} overwrite failed. hr=0x{:08X}", opName, static_cast<unsigned long>(opHr)));
        if (FAILED(opHr))
        {
            return false;
        }

        std::string finalDestinationPuid;
        readPersistentId(destinationPath, finalDestinationPuid, L"final destination");
        state.Require(! finalDestinationPuid.empty() && finalDestinationPuid != oldDestinationPuid,
                      std::format(L"MTP copy/move overwrite swap: {} destination did not take the temp object's persistent id. old={} final={}",
                                  opName,
                                  widenAscii(oldDestinationPuid),
                                  widenAscii(finalDestinationPuid)));

        std::string readBack;
        state.Require(ReadPluginFileText(io.get(), destinationPath.c_str(), readBack, state, L"MTP copy/move overwrite swap dest readback"),
                      std::format(L"MTP copy/move overwrite swap: {} destination readback failed.", opName));
        state.Require(readBack == kSourcePayload, std::format(L"MTP copy/move overwrite swap: {} destination contents did not match source.", opName));

        unsigned long sourceAttributes = 0;
        const HRESULT sourceAttrsHr    = io->GetAttributes(sourcePath.c_str(), &sourceAttributes);
        if (move)
        {
            state.Require(
                sourceAttrsHr == HRESULT_FROM_WIN32(ERROR_FILE_NOT_FOUND),
                std::format(L"MTP copy/move overwrite swap: move source should be deleted, got hr=0x{:08X}.", static_cast<unsigned long>(sourceAttrsHr)));
        }
        else
        {
            state.Require(SUCCEEDED(sourceAttrsHr),
                          std::format(L"MTP copy/move overwrite swap: copy source should remain, got hr=0x{:08X}.", static_cast<unsigned long>(sourceAttrsHr)));
        }

        const auto entries      = SnapshotDirectoryEntries(created.fileSystem, baseFolder.c_str(), state, L"MTP copy/move overwrite swap camera");
        uint32_t finalNameCount = 0;
        for (const DirectoryEntrySnapshot& entry : entries)
        {
            if (entry.name == destName)
            {
                ++finalNameCount;
            }
            state.Require(entry.name.find(L".rs-mtp-overwrite-") == std::wstring::npos,
                          std::format(L"MTP copy/move overwrite swap: leaked temp entry {} during {}.", entry.name, opName));
        }
        state.Require(finalNameCount == 1u, std::format(L"MTP copy/move overwrite swap: {} expected one final entry, got {}.", opName, finalNameCount));
        if (! state.failure.empty())
        {
            return false;
        }
    }

    return state.failure.empty();
});

SelfTest::RunCase(options,
                  suite,
                  L"mtp_overwrite_journal_clears_completed_swap_without_temp",
                  [&](SelfTest::CaseState& state) noexcept
{
    const std::wstring guid = MakeGuidText();
    state.Require(! guid.empty(), L"MTP journal completed swap: failed to generate unique file name.");
    if (guid.empty())
    {
        return false;
    }

    std::wstring previousLocalAppData;
    std::wstring localAppData;
    if (! AcquireMtpJournalLocalAppDataSandbox(
            state, L"mtp_journal_completed_swap", L"MTP journal completed swap", localAppData, previousLocalAppData))
    {
        return false;
    }
    const auto restoreLocalAppData = wil::scope_exit([&]() noexcept { RestoreMtpJournalLocalAppDataSandbox(previousLocalAppData); });

    constexpr std::wstring_view kBaseFolder = L"/Fake Phone [devpuid:fake-device]/Internal Storage/DCIM/Camera";
    const std::wstring host                 = L"journal-completed-swap-" + guid;
    const std::wstring finalName            = L"journal-completed-swap-" + guid + L".txt";
    const std::wstring tempName             = L"." + finalName + L".rs-mtp-overwrite-completed-swap-" + guid + L".tmp";
    const std::wstring finalPath            = std::wstring(kBaseFolder) + L"/" + finalName;
    const std::wstring tempPath             = std::wstring(kBaseFolder) + L"/" + tempName;
    const std::wstring journalDirectory     = localAppData + L"\\RedSalamander\\PluginState\\FileSystemMtp\\" + std::format(L"{:016X}", stableDeviceHash(host));
    const std::wstring journalPath          = journalDirectory + L"\\overwrite-journal.json";
    const std::string configuration         = std::format(R"json({{"readOnly":false,"host":"{}"}})json", narrowAscii(host));

    CreatedFileSystemInstance created;
    const HRESULT createHr = TryCreateFakeMtpFileSystemInstance("{}", configuration, L"/", created);
    if (createHr == HRESULT_FROM_WIN32(ERROR_PROC_NOT_FOUND))
    {
        return state.Skip(L"MTP fake backend export is available only in debug plugin builds.");
    }
    state.Require(SUCCEEDED(createHr) && created.fileSystem,
                  std::format(L"MTP journal completed swap: create selftest instance failed. hr=0x{:08X}", static_cast<unsigned long>(createHr)));
    if (FAILED(createHr) || ! created.fileSystem)
    {
        return false;
    }

    wil::com_ptr<IFileSystemIO> io;
    state.Require(CreateFileSystemIo(created.fileSystem, io), L"MTP journal completed swap: missing IFileSystemIO.");
    if (! io)
    {
        return false;
    }

    auto cleanup = wil::scope_exit([&]() noexcept
    {
        if (! journalPath.empty())
        {
            static_cast<void>(DeleteFileW(journalPath.c_str()));
        }
        static_cast<void>(DeleteFileW((journalPath + L".stale").c_str()));
        static_cast<void>(created.fileSystem->DeleteItem(finalPath.c_str(), FILESYSTEM_FLAG_NONE, nullptr, nullptr, nullptr));
        static_cast<void>(created.fileSystem->DeleteItem(tempPath.c_str(), FILESYSTEM_FLAG_NONE, nullptr, nullptr, nullptr));
    });

    const auto readPropertyBatchCalls = [&](uint64_t& value, std::wstring_view label) noexcept -> bool
    {
        value                  = 0;
        const char* properties = nullptr;
        const HRESULT propsHr  = io->GetItemProperties(finalPath.c_str(), &properties);
        state.Require(SUCCEEDED(propsHr) && properties != nullptr,
                      std::format(L"MTP journal completed swap: {} GetItemProperties failed. hr=0x{:08X}",
                                  label,
                                  static_cast<unsigned long>(propsHr)));
        if (FAILED(propsHr) || properties == nullptr)
        {
            return false;
        }

        const std::optional<uint64_t> parsed = extractJsonUInt(properties, "propertyBatchCalls");
        state.Require(parsed.has_value(), std::format(L"MTP journal completed swap: {} missing propertyBatchCalls.", label));
        if (! parsed.has_value())
        {
            return false;
        }

        value = parsed.value();
        return true;
    };

    constexpr std::string_view kFinalPayload = "replacement payload already committed before journal clear";
    state.Require(WritePluginFileText(io.get(), finalPath.c_str(), FILESYSTEM_FLAG_NONE, kFinalPayload, state, L"MTP journal completed swap final write"),
                  L"MTP journal completed swap: final write failed.");
    if (! state.failure.empty())
    {
        return false;
    }

    HRESULT hr = ensureDirectoryExists(journalDirectory);
    state.Require(SUCCEEDED(hr), std::format(L"MTP journal completed swap: ensure journal directory failed. hr=0x{:08X}", static_cast<unsigned long>(hr)));
    if (FAILED(hr))
    {
        return false;
    }

    const std::string journalJson = std::format(
        R"json({{"schemaVersion":1,"entries":[{{"phase":"committed","devicePuid":"{}","sourcePath":"","destinationPath":"{}","tempPath":"{}","declaredSizeBytes":{},"sourceTransmitHashHex":"manual-selftest","journalTimestampFileTimeUtc":1,"replayAttemptCount":0}}]}})json",
        narrowAscii(host),
        narrowAscii(finalPath),
        narrowAscii(tempPath),
        static_cast<unsigned long long>(kFinalPayload.size()));
    hr = writeUtf8File(journalPath, journalJson);
    state.Require(SUCCEEDED(hr), std::format(L"MTP journal completed swap: write journal failed. hr=0x{:08X}", static_cast<unsigned long>(hr)));
    if (FAILED(hr))
    {
        return false;
    }

    hr = notifyInjectedJournal(host);
    state.Require(SUCCEEDED(hr),
                  std::format(L"MTP journal completed swap: cache invalidation failed. hr=0x{:08X}", static_cast<unsigned long>(hr)));
    if (FAILED(hr))
    {
        return false;
    }

    std::string readBack;
    state.Require(ReadPluginFileText(io.get(), finalPath.c_str(), readBack, state, L"MTP journal completed swap first replay read"),
                  L"MTP journal completed swap: first replay read failed.");
    state.Require(readBack == kFinalPayload, L"MTP journal completed swap: completed final payload changed during replay.");

    const DWORD journalAttributes = GetFileAttributesW(journalPath.c_str());
    state.Require(journalAttributes == INVALID_FILE_ATTRIBUTES && GetLastError() == ERROR_FILE_NOT_FOUND,
                  L"MTP journal completed swap: replay retained the stale completed-swap journal.");

    uint64_t propertyBatchCallsAfterSecondCommand = 0;
    if (! readPropertyBatchCalls(propertyBatchCallsAfterSecondCommand, L"after second backend command"))
    {
        return false;
    }

    state.Require(propertyBatchCallsAfterSecondCommand <= 1u,
                  std::format(L"MTP journal completed swap: second backend command replayed stale journal. after={}.",
                              propertyBatchCallsAfterSecondCommand));

    return state.failure.empty();
});

SelfTest::RunCase(options,
                  suite,
                  L"mtp_rename_overwrite_uses_temp_puid_swap",
                  [&](SelfTest::CaseState& state) noexcept
{
    const auto extractJsonString = [](std::string_view json, std::string_view key) noexcept -> std::string
    {
        const std::string token = std::format(R"json("{}":")json", key);
        const size_t start      = json.find(token);
        if (start == std::string_view::npos)
        {
            return {};
        }

        const size_t valueStart = start + token.size();
        const size_t valueEnd   = json.find('"', valueStart);
        if (valueEnd == std::string_view::npos)
        {
            return {};
        }

        return std::string(json.substr(valueStart, valueEnd - valueStart));
    };
    const auto widenAscii = [](std::string_view text) { return std::wstring(text.begin(), text.end()); };

    CreatedFileSystemInstance created;
    const HRESULT createHr = TryCreateFakeMtpFileSystemInstance("{}", R"json({"readOnly":false})json", L"/", created);
    if (createHr == HRESULT_FROM_WIN32(ERROR_PROC_NOT_FOUND))
    {
        return state.Skip(L"MTP fake backend export is available only in debug plugin builds.");
    }
    state.Require(SUCCEEDED(createHr) && created.fileSystem,
                  std::format(L"MTP rename overwrite swap: create selftest instance failed. hr=0x{:08X}", static_cast<unsigned long>(createHr)));
    if (FAILED(createHr) || ! created.fileSystem)
    {
        return false;
    }

    wil::com_ptr<IFileSystemIO> io;
    state.Require(CreateFileSystemIo(created.fileSystem, io), L"MTP rename overwrite swap: missing IFileSystemIO.");
    if (! io)
    {
        return false;
    }

    const std::wstring guid = MakeGuidText();
    state.Require(! guid.empty(), L"MTP rename overwrite swap: failed to generate unique file name.");
    if (guid.empty())
    {
        return false;
    }

    const std::wstring baseFolder      = L"/Fake Phone [devpuid:fake-device]/Internal Storage/DCIM/Camera";
    const std::wstring sourcePath      = baseFolder + L"/rename-swap-" + guid + L"-source.txt";
    const std::wstring destName        = L"rename-swap-" + guid + L"-dest.txt";
    const std::wstring destinationPath = baseFolder + L"/" + destName;
    auto cleanup                       = wil::scope_exit([&]() noexcept
    {
        static_cast<void>(created.fileSystem->DeleteItem(sourcePath.c_str(), FILESYSTEM_FLAG_NONE, nullptr, nullptr, nullptr));
        static_cast<void>(created.fileSystem->DeleteItem(destinationPath.c_str(), FILESYSTEM_FLAG_NONE, nullptr, nullptr, nullptr));
    });

    constexpr std::string_view kSourcePayload = "rename source payload that should become destination";
    constexpr std::string_view kOldPayload    = "rename destination payload to be replaced";
    state.Require(WritePluginFileText(io.get(), sourcePath.c_str(), FILESYSTEM_FLAG_NONE, kSourcePayload, state, L"MTP rename overwrite swap source write"),
                  L"MTP rename overwrite swap: source write failed.");
    state.Require(WritePluginFileText(io.get(), destinationPath.c_str(), FILESYSTEM_FLAG_NONE, kOldPayload, state, L"MTP rename overwrite swap dest write"),
                  L"MTP rename overwrite swap: destination write failed.");
    if (! state.failure.empty())
    {
        return false;
    }

    const auto readPersistentId = [&](const std::wstring& path, std::string& persistentId, std::wstring_view label) noexcept
    {
        persistentId.clear();
        const char* properties = nullptr;
        const HRESULT propsHr  = io->GetItemProperties(path.c_str(), &properties);
        state.Require(SUCCEEDED(propsHr) && properties != nullptr,
                      std::format(L"MTP rename overwrite swap: {} GetItemProperties failed. hr=0x{:08X}", label, static_cast<unsigned long>(propsHr)));
        if (FAILED(propsHr) || properties == nullptr)
        {
            return;
        }

        persistentId = extractJsonString(properties, "persistentId");
        state.Require(! persistentId.empty(), std::format(L"MTP rename overwrite swap: {} did not expose a persistent id.", label));
    };

    std::string sourcePuid;
    std::string oldDestinationPuid;
    readPersistentId(sourcePath, sourcePuid, L"source");
    readPersistentId(destinationPath, oldDestinationPuid, L"old destination");
    if (! state.failure.empty())
    {
        return false;
    }

    const HRESULT renameHr =
        created.fileSystem->RenameItem(sourcePath.c_str(), destinationPath.c_str(), FILESYSTEM_FLAG_ALLOW_OVERWRITE, nullptr, nullptr, nullptr);
    state.Require(SUCCEEDED(renameHr), std::format(L"MTP rename overwrite swap: RenameItem failed. hr=0x{:08X}", static_cast<unsigned long>(renameHr)));
    if (FAILED(renameHr))
    {
        return false;
    }

    const char* finalProperties = nullptr;
    const HRESULT propsHr       = io->GetItemProperties(destinationPath.c_str(), &finalProperties);
    state.Require(SUCCEEDED(propsHr) && finalProperties != nullptr,
                  std::format(L"MTP rename overwrite swap: final GetItemProperties failed. hr=0x{:08X}", static_cast<unsigned long>(propsHr)));
    if (FAILED(propsHr) || finalProperties == nullptr)
    {
        return false;
    }

    const std::string_view props(finalProperties);
    const std::string finalDestinationPuid = extractJsonString(props, "persistentId");
    state.Require(! finalDestinationPuid.empty() && finalDestinationPuid != oldDestinationPuid,
                  std::format(L"MTP rename overwrite swap: destination did not replace the old object. old={} final={}",
                              widenAscii(oldDestinationPuid),
                              widenAscii(finalDestinationPuid)));
    state.Require(! finalDestinationPuid.empty() && finalDestinationPuid != sourcePuid,
                  std::format(L"MTP rename overwrite swap: destination kept the source PUID instead of taking a temp PUID. source={} final={}",
                              widenAscii(sourcePuid),
                              widenAscii(finalDestinationPuid)));
    state.Require(extractJsonUInt(props, "copyItemCalls") == 1u,
                  L"MTP rename overwrite swap: fake backend trace did not show the temp-copy overwrite path.");

    std::string readBack;
    state.Require(ReadPluginFileText(io.get(), destinationPath.c_str(), readBack, state, L"MTP rename overwrite swap dest readback"),
                  L"MTP rename overwrite swap: destination readback failed.");
    state.Require(readBack == kSourcePayload, L"MTP rename overwrite swap: destination contents did not match source.");

    unsigned long sourceAttributes = 0;
    const HRESULT sourceAttrsHr    = io->GetAttributes(sourcePath.c_str(), &sourceAttributes);
    state.Require(sourceAttrsHr == HRESULT_FROM_WIN32(ERROR_FILE_NOT_FOUND),
                  std::format(L"MTP rename overwrite swap: source should be deleted, got hr=0x{:08X}.", static_cast<unsigned long>(sourceAttrsHr)));

    const auto entries      = SnapshotDirectoryEntries(created.fileSystem, baseFolder.c_str(), state, L"MTP rename overwrite swap camera");
    uint32_t finalNameCount = 0;
    for (const DirectoryEntrySnapshot& entry : entries)
    {
        if (entry.name == destName)
        {
            ++finalNameCount;
        }
        state.Require(entry.name.find(L".rs-mtp-overwrite-") == std::wstring::npos,
                      std::format(L"MTP rename overwrite swap: leaked temp entry {}.", entry.name));
    }
    state.Require(finalNameCount == 1u, std::format(L"MTP rename overwrite swap: expected one final entry, got {}.", finalNameCount));

    return state.failure.empty();
});

SelfTest::RunCase(options,
                  suite,
                  L"mtp_copy_move_overwrite_temp_copy_failure_keeps_original_and_allows_retry",
                  [&](SelfTest::CaseState& state) noexcept
{
    const std::wstring guid = MakeGuidText();
    state.Require(! guid.empty(), L"MTP copy/move temp-copy failure: failed to generate unique file name.");
    if (guid.empty())
    {
        return false;
    }

    const std::wstring baseFolder             = L"/Fake Phone [devpuid:fake-device]/Internal Storage/DCIM/Camera";
    constexpr std::array<bool, 2> kMoveValues = {{false, true}};
    for (const bool move : kMoveValues)
    {
        CreatedFileSystemInstance created;
        const HRESULT createHr = TryCreateFakeMtpFileSystemInstance(
            R"json({"copyItemFailOnceDestinationPathContains":".rs-mtp-overwrite-"})json", R"json({"readOnly":false})json", L"/", created);
        if (createHr == HRESULT_FROM_WIN32(ERROR_PROC_NOT_FOUND))
        {
            return state.Skip(L"MTP fake backend export is available only in debug plugin builds.");
        }
        state.Require(SUCCEEDED(createHr) && created.fileSystem,
                      std::format(L"MTP copy/move temp-copy failure: create selftest instance failed. hr=0x{:08X}", static_cast<unsigned long>(createHr)));
        if (FAILED(createHr) || ! created.fileSystem)
        {
            return false;
        }

        wil::com_ptr<IFileSystemIO> io;
        state.Require(CreateFileSystemIo(created.fileSystem, io), L"MTP copy/move temp-copy failure: missing IFileSystemIO.");
        if (! io)
        {
            return false;
        }

        const std::wstring opName          = move ? L"move" : L"copy";
        const std::wstring sourcePath      = baseFolder + L"/temp-copy-failure-" + opName + L"-" + guid + L"-source.txt";
        const std::wstring destName        = L"temp-copy-failure-" + opName + L"-" + guid + L"-dest.txt";
        const std::wstring destinationPath = baseFolder + L"/" + destName;
        auto cleanup                       = wil::scope_exit([&]() noexcept
        {
            static_cast<void>(created.fileSystem->DeleteItem(sourcePath.c_str(), FILESYSTEM_FLAG_NONE, nullptr, nullptr, nullptr));
            static_cast<void>(created.fileSystem->DeleteItem(destinationPath.c_str(), FILESYSTEM_FLAG_NONE, nullptr, nullptr, nullptr));
        });

        const auto requireSingleFinalAndNoTemp = [&](std::wstring_view label) noexcept
        {
            const auto entries      = SnapshotDirectoryEntries(created.fileSystem, baseFolder.c_str(), state, L"MTP copy/move temp-copy failure camera");
            uint32_t finalNameCount = 0;
            for (const DirectoryEntrySnapshot& entry : entries)
            {
                if (entry.name == destName)
                {
                    ++finalNameCount;
                }
                state.Require(entry.name.find(L".rs-mtp-overwrite-") == std::wstring::npos,
                              std::format(L"MTP copy/move temp-copy failure: leaked temp entry {} during {} {}.", entry.name, opName, label));
            }
            state.Require(finalNameCount == 1u,
                          std::format(L"MTP copy/move temp-copy failure: {} expected one final entry during {}, got {}.", opName, label, finalNameCount));
        };

        constexpr std::string_view kSourcePayload = "copy/move temp-copy failure source payload";
        constexpr std::string_view kOldPayload    = "copy/move temp-copy failure old destination";
        state.Require(
            WritePluginFileText(io.get(), sourcePath.c_str(), FILESYSTEM_FLAG_NONE, kSourcePayload, state, L"MTP copy/move temp-copy failure source write"),
            std::format(L"MTP copy/move temp-copy failure: {} source write failed.", opName));
        state.Require(
            WritePluginFileText(io.get(), destinationPath.c_str(), FILESYSTEM_FLAG_NONE, kOldPayload, state, L"MTP copy/move temp-copy failure dest write"),
            std::format(L"MTP copy/move temp-copy failure: {} destination write failed.", opName));
        if (! state.failure.empty())
        {
            return false;
        }

        const auto runOverwrite = [&]() noexcept -> HRESULT
        {
            return move ? created.fileSystem->MoveItem(sourcePath.c_str(), destinationPath.c_str(), FILESYSTEM_FLAG_ALLOW_OVERWRITE, nullptr, nullptr, nullptr)
                        : created.fileSystem->CopyItem(sourcePath.c_str(), destinationPath.c_str(), FILESYSTEM_FLAG_ALLOW_OVERWRITE, nullptr, nullptr, nullptr);
        };

        const HRESULT failedHr = runOverwrite();
        state.Require(
            failedHr == HRESULT_FROM_WIN32(ERROR_WRITE_FAULT),
            std::format(L"MTP copy/move temp-copy failure: {} expected ERROR_WRITE_FAULT, got hr=0x{:08X}.", opName, static_cast<unsigned long>(failedHr)));
        if (! state.failure.empty())
        {
            return false;
        }

        std::string sourceReadBack;
        std::string destReadBack;
        state.Require(ReadPluginFileText(io.get(), sourcePath.c_str(), sourceReadBack, state, L"MTP copy/move temp-copy failure source after fail"),
                      std::format(L"MTP copy/move temp-copy failure: {} source readback after failure failed.", opName));
        state.Require(ReadPluginFileText(io.get(), destinationPath.c_str(), destReadBack, state, L"MTP copy/move temp-copy failure dest after fail"),
                      std::format(L"MTP copy/move temp-copy failure: {} destination readback after failure failed.", opName));
        state.Require(sourceReadBack == kSourcePayload, std::format(L"MTP copy/move temp-copy failure: {} failure changed source contents.", opName));
        state.Require(destReadBack == kOldPayload, std::format(L"MTP copy/move temp-copy failure: {} failure changed destination contents.", opName));
        requireSingleFinalAndNoTemp(L"after failed temp copy");
        if (! state.failure.empty())
        {
            return false;
        }

        const HRESULT retryHr = runOverwrite();
        state.Require(SUCCEEDED(retryHr),
                      std::format(L"MTP copy/move temp-copy failure: {} retry failed. hr=0x{:08X}", opName, static_cast<unsigned long>(retryHr)));
        if (! state.failure.empty())
        {
            return false;
        }

        destReadBack.clear();
        state.Require(ReadPluginFileText(io.get(), destinationPath.c_str(), destReadBack, state, L"MTP copy/move temp-copy failure dest retry"),
                      std::format(L"MTP copy/move temp-copy failure: {} destination retry readback failed.", opName));
        state.Require(destReadBack == kSourcePayload,
                      std::format(L"MTP copy/move temp-copy failure: {} retry destination contents did not match source.", opName));

        unsigned long sourceAttributes = 0;
        const HRESULT sourceAttrsHr    = io->GetAttributes(sourcePath.c_str(), &sourceAttributes);
        if (move)
        {
            state.Require(
                sourceAttrsHr == HRESULT_FROM_WIN32(ERROR_FILE_NOT_FOUND),
                std::format(L"MTP copy/move temp-copy failure: move retry should delete source, got hr=0x{:08X}.", static_cast<unsigned long>(sourceAttrsHr)));
        }
        else
        {
            state.Require(
                SUCCEEDED(sourceAttrsHr),
                std::format(L"MTP copy/move temp-copy failure: copy retry should keep source, got hr=0x{:08X}.", static_cast<unsigned long>(sourceAttrsHr)));
        }

        requireSingleFinalAndNoTemp(L"after retry");
        if (! state.failure.empty())
        {
            return false;
        }
    }

    return state.failure.empty();
});

SelfTest::RunCase(options,
                  suite,
                  L"mtp_move_fallback_delete_source_failure_leaves_duplicate_and_reports_partial",
                  [&](SelfTest::CaseState& state) noexcept
{
    CreatedFileSystemInstance created;
    const HRESULT createHr =
        TryCreateFakeMtpFileSystemInstance(R"json({"moveFallbackDeleteSourceFails":true})json", R"json({"readOnly":false})json", L"/", created);
    if (createHr == HRESULT_FROM_WIN32(ERROR_PROC_NOT_FOUND))
    {
        return state.Skip(L"MTP fake backend export is available only in debug plugin builds.");
    }
    state.Require(SUCCEEDED(createHr) && created.fileSystem,
                  std::format(L"MTP move fallback partial: create selftest instance failed. hr=0x{:08X}", static_cast<unsigned long>(createHr)));
    if (FAILED(createHr) || ! created.fileSystem)
    {
        return false;
    }

    wil::com_ptr<IFileSystemIO> io;
    wil::com_ptr<IFileSystemDirectoryOperations> dirOps;
    state.Require(CreateFileSystemIo(created.fileSystem, io), L"MTP move fallback partial: missing IFileSystemIO.");
    state.Require(CreateFileSystemDirectoryOperations(created.fileSystem, dirOps), L"MTP move fallback partial: missing IFileSystemDirectoryOperations.");
    if (! io || ! dirOps)
    {
        return false;
    }

    const std::wstring guid = MakeGuidText();
    state.Require(! guid.empty(), L"MTP move fallback partial: failed to generate unique case name.");
    if (guid.empty())
    {
        return false;
    }

    const std::wstring caseRoot = std::format(L"/Fake Phone [devpuid:fake-device]/Internal Storage/DCIM/Camera/move-fallback-{}", guid);
    const HRESULT mkdirHr       = dirOps->CreateDirectory(caseRoot.c_str());
    state.Require(SUCCEEDED(mkdirHr), std::format(L"MTP move fallback partial: CreateDirectory failed. hr=0x{:08X}", static_cast<unsigned long>(mkdirHr)));
    if (FAILED(mkdirHr))
    {
        return false;
    }

    auto cleanup = wil::scope_exit([&]() noexcept
    { static_cast<void>(created.fileSystem->DeleteItem(caseRoot.c_str(), FILESYSTEM_FLAG_RECURSIVE, nullptr, nullptr, nullptr)); });

    constexpr std::string_view kPayload = "move fallback delete-source failure payload";
    const std::wstring sourcePath       = caseRoot + L"/source.txt";
    const std::wstring destinationPath  = caseRoot + L"/destination.txt";
    state.Require(WritePluginFileText(io.get(), sourcePath.c_str(), FILESYSTEM_FLAG_NONE, kPayload, state, L"MTP move fallback partial write"),
                  L"MTP move fallback partial: source write failed.");
    if (! state.failure.empty())
    {
        return false;
    }

    RecordingFileSystemCallback callback(1);
    const HRESULT moveHr = created.fileSystem->MoveItem(sourcePath.c_str(), destinationPath.c_str(), FILESYSTEM_FLAG_NONE, nullptr, &callback, nullptr);
    constexpr HRESULT kExpectedFailure = HRESULT_FROM_WIN32(ERROR_ACCESS_DENIED);
    state.Require(moveHr == kExpectedFailure,
                  std::format(L"MTP move fallback partial: MoveItem expected delete-source failure 0x{:08X}, got 0x{:08X}.",
                              static_cast<unsigned long>(kExpectedFailure),
                              static_cast<unsigned long>(moveHr)));
    state.Require(callback.CompletedCount() == 1u,
                  std::format(L"MTP move fallback partial: expected 1 completion callback, got {}.", callback.CompletedCount()));
    state.Require(! callback.SawUnexpectedIssue(), L"MTP move fallback partial: unexpected issue callback was raised.");

    RecordedFileSystemItem completed{};
    state.Require(callback.TryGetItem(0, completed), L"MTP move fallback partial: completion callback was not recorded.");
    if (completed.seen)
    {
        state.Require(completed.status == kExpectedFailure,
                      std::format(L"MTP move fallback partial: completion status expected 0x{:08X}, got 0x{:08X}.",
                                  static_cast<unsigned long>(kExpectedFailure),
                                  static_cast<unsigned long>(completed.status)));
        state.Require(completed.sourcePath == sourcePath, L"MTP move fallback partial: completion source path did not match.");
        state.Require(completed.destinationPath == destinationPath, L"MTP move fallback partial: completion destination path did not match.");
    }

    std::string sourceReadBack;
    std::string destinationReadBack;
    state.Require(ReadPluginFileText(io.get(), sourcePath.c_str(), sourceReadBack, state, L"MTP move fallback partial source read"),
                  L"MTP move fallback partial: source was not left readable after failed delete-source phase.");
    state.Require(ReadPluginFileText(io.get(), destinationPath.c_str(), destinationReadBack, state, L"MTP move fallback partial destination read"),
                  L"MTP move fallback partial: destination copy was not left readable after failed delete-source phase.");
    state.Require(sourceReadBack == kPayload, L"MTP move fallback partial: source payload changed.");
    state.Require(destinationReadBack == kPayload, L"MTP move fallback partial: destination payload did not match the source.");

    return state.failure.empty();
});

SelfTest::RunCase(options,
                  suite,
                  L"mtp_transfer_cancel_is_prompt",
                  [&](SelfTest::CaseState& state) noexcept
{
    class CancelOnFirstTransferCallback final : public IFileSystemCallback
    {
    public:
        CancelOnFirstTransferCallback() = default;

        CancelOnFirstTransferCallback(const CancelOnFirstTransferCallback&)            = delete;
        CancelOnFirstTransferCallback(CancelOnFirstTransferCallback&&)                 = delete;
        CancelOnFirstTransferCallback& operator=(const CancelOnFirstTransferCallback&) = delete;
        CancelOnFirstTransferCallback& operator=(CancelOnFirstTransferCallback&&)      = delete;

        HRESULT STDMETHODCALLTYPE FileSystemProgress([[maybe_unused]] FileSystemOperation operationType,
                                                     [[maybe_unused]] unsigned long totalItems,
                                                     [[maybe_unused]] unsigned long completedItems,
                                                     [[maybe_unused]] uint64_t totalBytes,
                                                     [[maybe_unused]] uint64_t completedBytes,
                                                     [[maybe_unused]] const wchar_t* currentSourcePath,
                                                     [[maybe_unused]] const wchar_t* currentDestinationPath,
                                                     [[maybe_unused]] uint64_t currentItemTotalBytes,
                                                     [[maybe_unused]] uint64_t currentItemCompletedBytes,
                                                     [[maybe_unused]] FileSystemOptions* options,
                                                     [[maybe_unused]] uint64_t progressStreamId,
                                                     [[maybe_unused]] void* cookie) noexcept override
        {
            _progressCalls.fetch_add(1u, std::memory_order_acq_rel);
            return S_OK;
        }

        HRESULT STDMETHODCALLTYPE FileSystemItemCompleted([[maybe_unused]] FileSystemOperation operationType,
                                                          [[maybe_unused]] unsigned long itemIndex,
                                                          [[maybe_unused]] const wchar_t* sourcePath,
                                                          [[maybe_unused]] const wchar_t* destinationPath,
                                                          HRESULT status,
                                                          [[maybe_unused]] FileSystemOptions* options,
                                                          [[maybe_unused]] void* cookie) noexcept override
        {
            _completionStatus.store(status, std::memory_order_release);
            _completedCalls.fetch_add(1u, std::memory_order_acq_rel);
            return S_OK;
        }

        HRESULT STDMETHODCALLTYPE FileSystemShouldCancel(BOOL* pCancel, [[maybe_unused]] void* cookie) noexcept override
        {
            if (pCancel == nullptr)
            {
                return E_POINTER;
            }

            _cancelChecks.fetch_add(1u, std::memory_order_acq_rel);
            *pCancel = TRUE;
            return S_OK;
        }

        HRESULT STDMETHODCALLTYPE FileSystemIssue([[maybe_unused]] FileSystemOperation operationType,
                                                  [[maybe_unused]] const wchar_t* sourcePath,
                                                  [[maybe_unused]] const wchar_t* destinationPath,
                                                  [[maybe_unused]] HRESULT status,
                                                  FileSystemIssueAction* action,
                                                  [[maybe_unused]] FileSystemOptions* options,
                                                  [[maybe_unused]] void* cookie) noexcept override
        {
            _issues.fetch_add(1u, std::memory_order_acq_rel);
            if (action)
            {
                *action = FileSystemIssueAction::Cancel;
            }
            return E_UNEXPECTED;
        }

        [[nodiscard]] uint32_t ProgressCalls() const noexcept
        {
            return _progressCalls.load(std::memory_order_acquire);
        }

        [[nodiscard]] uint32_t CancelChecks() const noexcept
        {
            return _cancelChecks.load(std::memory_order_acquire);
        }

        [[nodiscard]] uint32_t CompletedCalls() const noexcept
        {
            return _completedCalls.load(std::memory_order_acquire);
        }

        [[nodiscard]] HRESULT CompletionStatus() const noexcept
        {
            return _completionStatus.load(std::memory_order_acquire);
        }

        [[nodiscard]] uint32_t IssueCalls() const noexcept
        {
            return _issues.load(std::memory_order_acquire);
        }

    private:
        std::atomic_uint32_t _progressCalls{0};
        std::atomic_uint32_t _cancelChecks{0};
        std::atomic_uint32_t _completedCalls{0};
        std::atomic_uint32_t _issues{0};
        std::atomic<HRESULT> _completionStatus{S_OK};
    };

    CreatedFileSystemInstance created;
    const HRESULT createHr = TryCreateFakeMtpFileSystemInstance(R"json({"operationDelayMs":1000})json", R"json({"readOnly":false})json", L"/", created);
    if (createHr == HRESULT_FROM_WIN32(ERROR_PROC_NOT_FOUND))
    {
        return state.Skip(L"MTP fake backend export is available only in debug plugin builds.");
    }
    state.Require(SUCCEEDED(createHr) && created.fileSystem,
                  std::format(L"MTP transfer cancel: create selftest instance failed. hr=0x{:08X}", static_cast<unsigned long>(createHr)));
    if (FAILED(createHr) || ! created.fileSystem)
    {
        return false;
    }

    wil::com_ptr<IFileSystemIO> io;
    state.Require(CreateFileSystemIo(created.fileSystem, io), L"MTP transfer cancel: missing IFileSystemIO.");
    if (! io)
    {
        return false;
    }

    const std::wstring guid = MakeGuidText();
    state.Require(! guid.empty(), L"MTP transfer cancel: failed to generate unique case name.");
    if (guid.empty())
    {
        return false;
    }

    const std::wstring sourcePath      = L"/Fake Phone [devpuid:fake-device]/Internal Storage/DCIM/Camera/photo001.txt";
    const std::wstring destinationPath = L"/Fake Phone [devpuid:fake-device]/Internal Storage/DCIM/Camera/cancel-destination-" + guid + L".txt";

    CancelOnFirstTransferCallback callback;
    const auto start     = std::chrono::steady_clock::now();
    const HRESULT copyHr = created.fileSystem->CopyItem(sourcePath.c_str(), destinationPath.c_str(), FILESYSTEM_FLAG_NONE, nullptr, &callback, nullptr);
    const auto elapsedMs = std::chrono::duration_cast<std::chrono::milliseconds>(std::chrono::steady_clock::now() - start).count();
    constexpr HRESULT kExpectedCancel = HRESULT_FROM_WIN32(ERROR_CANCELLED);
    state.Require(copyHr == kExpectedCancel,
                  std::format(L"MTP transfer cancel: CopyItem expected ERROR_CANCELLED, got 0x{:08X}.", static_cast<unsigned long>(copyHr)));
    state.Require(elapsedMs < 500, std::format(L"MTP transfer cancel: cancellation took {} ms; expected prompt return before backend delay.", elapsedMs));
    state.Require(callback.ProgressCalls() >= 1u, L"MTP transfer cancel: no progress callback was emitted before cancellation polling.");
    state.Require(callback.CancelChecks() >= 1u, L"MTP transfer cancel: FileSystemShouldCancel was not polled.");
    state.Require(callback.CompletedCalls() == 1u, std::format(L"MTP transfer cancel: expected 1 completion callback, got {}.", callback.CompletedCalls()));
    state.Require(callback.CompletionStatus() == kExpectedCancel,
                  std::format(L"MTP transfer cancel: completion status expected 0x{:08X}, got 0x{:08X}.",
                              static_cast<unsigned long>(kExpectedCancel),
                              static_cast<unsigned long>(callback.CompletionStatus())));
    state.Require(callback.IssueCalls() == 0u, L"MTP transfer cancel: unexpected issue callback was raised.");

    unsigned long attrs       = 0;
    const HRESULT destAttrsHr = io->GetAttributes(destinationPath.c_str(), &attrs);
    state.Require(destAttrsHr == HRESULT_FROM_WIN32(ERROR_FILE_NOT_FOUND), L"MTP transfer cancel: cancelled copy created a destination item.");
    if (SUCCEEDED(destAttrsHr))
    {
        static_cast<void>(created.fileSystem->DeleteItem(destinationPath.c_str(), FILESYSTEM_FLAG_NONE, nullptr, nullptr, nullptr));
    }

    return state.failure.empty();
});

SelfTest::RunCase(options,
                  suite,
                  L"mtp_batch_callbacks_report_item_indices",
                  [&](SelfTest::CaseState& state) noexcept
{
    CreatedFileSystemInstance created;
    const HRESULT createHr = TryCreateFakeMtpFileSystemInstance("{}", R"json({"readOnly":false})json", L"/", created);
    if (createHr == HRESULT_FROM_WIN32(ERROR_PROC_NOT_FOUND))
    {
        return state.Skip(L"MTP fake backend export is available only in debug plugin builds.");
    }
    state.Require(SUCCEEDED(createHr) && created.fileSystem,
                  std::format(L"MTP batch callback indices: create selftest instance failed. hr=0x{:08X}", static_cast<unsigned long>(createHr)));
    if (FAILED(createHr) || ! created.fileSystem)
    {
        return false;
    }

    wil::com_ptr<IFileSystemDirectoryOperations> dirOps;
    state.Require(CreateFileSystemDirectoryOperations(created.fileSystem, dirOps), L"MTP batch callback indices: missing IFileSystemDirectoryOperations.");
    if (! dirOps)
    {
        return false;
    }

    const std::wstring guid = MakeGuidText();
    state.Require(! guid.empty(), L"MTP batch callback indices: failed to generate unique case name.");
    if (guid.empty())
    {
        return false;
    }

    const std::wstring cameraPath        = L"/Fake Phone [devpuid:fake-device]/Internal Storage/DCIM/Camera";
    const std::wstring destinationFolder = cameraPath + L"/batch-callback-" + guid;
    const HRESULT mkdirHr                = dirOps->CreateDirectory(destinationFolder.c_str());
    state.Require(SUCCEEDED(mkdirHr), std::format(L"MTP batch callback indices: CreateDirectory failed. hr=0x{:08X}", static_cast<unsigned long>(mkdirHr)));
    if (FAILED(mkdirHr))
    {
        return false;
    }

    auto cleanup = wil::scope_exit([&]() noexcept
    { static_cast<void>(created.fileSystem->DeleteItem(destinationFolder.c_str(), FILESYSTEM_FLAG_RECURSIVE, nullptr, nullptr, nullptr)); });

    const std::wstring firstSource       = cameraPath + L"/photo001.txt";
    const std::wstring secondSource      = cameraPath + L"/name [puid:literal].txt";
    const std::wstring firstDestination  = destinationFolder + L"/photo001.txt";
    const std::wstring secondDestination = destinationFolder + L"/name [puid:literal].txt";
    const wchar_t* sources[]             = {firstSource.c_str(), secondSource.c_str()};

    RecordingFileSystemCallback callback(2);
    const HRESULT copyHr = created.fileSystem->CopyItems(sources, 2u, destinationFolder.c_str(), FILESYSTEM_FLAG_NONE, nullptr, &callback, nullptr);
    state.Require(SUCCEEDED(copyHr), std::format(L"MTP batch callback indices: CopyItems failed. hr=0x{:08X}", static_cast<unsigned long>(copyHr)));
    state.Require(callback.CompletedCount() == 2u,
                  std::format(L"MTP batch callback indices: expected 2 completion callbacks, got {}.", callback.CompletedCount()));
    state.Require(! callback.SawUnexpectedIssue(), L"MTP batch callback indices: unexpected issue callback was raised.");

    RecordedFileSystemItem firstCompleted{};
    RecordedFileSystemItem secondCompleted{};
    state.Require(callback.TryGetItem(0, firstCompleted), L"MTP batch callback indices: item 0 completion was not recorded.");
    state.Require(callback.TryGetItem(1, secondCompleted), L"MTP batch callback indices: item 1 completion was not recorded.");
    if (firstCompleted.seen)
    {
        state.Require(SUCCEEDED(firstCompleted.status),
                      std::format(L"MTP batch callback indices: item 0 failed. hr=0x{:08X}", static_cast<unsigned long>(firstCompleted.status)));
        state.Require(firstCompleted.sourcePath == firstSource, L"MTP batch callback indices: item 0 source path mismatch.");
        state.Require(firstCompleted.destinationPath == firstDestination, L"MTP batch callback indices: item 0 destination path mismatch.");
    }
    if (secondCompleted.seen)
    {
        state.Require(SUCCEEDED(secondCompleted.status),
                      std::format(L"MTP batch callback indices: item 1 failed. hr=0x{:08X}", static_cast<unsigned long>(secondCompleted.status)));
        state.Require(secondCompleted.sourcePath == secondSource, L"MTP batch callback indices: item 1 source path mismatch.");
        state.Require(secondCompleted.destinationPath == secondDestination, L"MTP batch callback indices: item 1 destination path mismatch.");
    }

    return state.failure.empty();
});

SelfTest::RunCase(options,
                  suite,
                  L"mtp_copy_from_device_accounting",
                  [&](SelfTest::CaseState& state) noexcept
{
    CreatedFileSystemInstance created;
    const HRESULT createHr = TryCreateFakeMtpFileSystemInstance("{}", R"json({"readOnly":false})json", L"/", created);
    if (createHr == HRESULT_FROM_WIN32(ERROR_PROC_NOT_FOUND))
    {
        return state.Skip(L"MTP fake backend export is available only in debug plugin builds.");
    }
    state.Require(SUCCEEDED(createHr) && created.fileSystem,
                  std::format(L"MTP copy from device accounting: create selftest instance failed. hr=0x{:08X}", static_cast<unsigned long>(createHr)));
    if (FAILED(createHr) || ! created.fileSystem)
    {
        return false;
    }

    wil::com_ptr<IFileSystemIO> io;
    state.Require(CreateFileSystemIo(created.fileSystem, io), L"MTP copy from device accounting: missing IFileSystemIO.");
    if (! io)
    {
        return false;
    }

    constexpr std::wstring_view kPhotoPath = L"/Fake Phone [devpuid:fake-device]/Internal Storage/DCIM/Camera/photo001.txt";
    std::string readBack;
    state.Require(ReadPluginFileText(io.get(), kPhotoPath.data(), readBack, state, L"MTP copy from device accounting read"),
                  L"MTP copy from device accounting: failed to read fixture.");
    state.Require(readBack == "RedSalamander deterministic MTP fixture\r\n", L"MTP copy from device accounting: read payload did not match.");

    const char* properties = nullptr;
    const HRESULT propsHr  = io->GetItemProperties(kPhotoPath.data(), &properties);
    state.Require(SUCCEEDED(propsHr) && properties != nullptr,
                  std::format(L"MTP copy from device accounting: GetItemProperties failed. hr=0x{:08X}", static_cast<unsigned long>(propsHr)));
    if (SUCCEEDED(propsHr) && properties)
    {
        const std::string_view props(properties);
        state.Require(props.find(R"json("readFileCalls":1)json") != std::string_view::npos,
                      L"MTP copy from device accounting: fake backend did not record exactly one read call.");
        state.Require(props.find(R"json("lastReadBytes":41)json") != std::string_view::npos,
                      L"MTP copy from device accounting: fake backend did not record the read byte count.");
    }

    return state.failure.empty();
});

SelfTest::RunCase(options,
                  suite,
                  L"mtp_copy_to_device_accounting",
                  [&](SelfTest::CaseState& state) noexcept
{
    CreatedFileSystemInstance created;
    const HRESULT createHr = TryCreateFakeMtpFileSystemInstance("{}", R"json({"readOnly":false})json", L"/", created);
    if (createHr == HRESULT_FROM_WIN32(ERROR_PROC_NOT_FOUND))
    {
        return state.Skip(L"MTP fake backend export is available only in debug plugin builds.");
    }
    state.Require(SUCCEEDED(createHr) && created.fileSystem,
                  std::format(L"MTP copy accounting: create selftest instance failed. hr=0x{:08X}", static_cast<unsigned long>(createHr)));
    if (FAILED(createHr) || ! created.fileSystem)
    {
        return false;
    }

    wil::com_ptr<IFileSystemIO> io;
    state.Require(CreateFileSystemIo(created.fileSystem, io), L"MTP copy accounting: missing IFileSystemIO.");
    if (! io)
    {
        return false;
    }

    const std::wstring guid = MakeGuidText();
    state.Require(! guid.empty(), L"MTP copy accounting: failed to generate unique case name.");
    if (guid.empty())
    {
        return false;
    }

    const std::wstring sourcePath      = L"/Fake Phone [devpuid:fake-device]/Internal Storage/DCIM/Camera/photo001.txt";
    const std::wstring destinationPath = L"/Fake Phone [devpuid:fake-device]/Internal Storage/DCIM/Camera/copy-accounting-" + guid + L".txt";
    auto cleanup                       = wil::scope_exit([&]() noexcept
    { static_cast<void>(created.fileSystem->DeleteItem(destinationPath.c_str(), FILESYSTEM_FLAG_NONE, nullptr, nullptr, nullptr)); });

    RecordingFileSystemCallback callback(1);
    const HRESULT copyHr = created.fileSystem->CopyItem(sourcePath.c_str(), destinationPath.c_str(), FILESYSTEM_FLAG_NONE, nullptr, &callback, nullptr);
    state.Require(SUCCEEDED(copyHr), std::format(L"MTP copy accounting: CopyItem failed. hr=0x{:08X}", static_cast<unsigned long>(copyHr)));
    state.Require(callback.ProgressCount() >= 1u, L"MTP copy accounting: expected at least one progress callback.");
    state.Require(callback.CompletedCount() == 1u, std::format(L"MTP copy accounting: expected 1 completion callback, got {}.", callback.CompletedCount()));
    state.Require(! callback.SawUnexpectedIssue(), L"MTP copy accounting: unexpected issue callback was raised.");

    RecordedFileSystemItem completed{};
    state.Require(callback.TryGetItem(0, completed) && SUCCEEDED(completed.status), L"MTP copy accounting: completion callback did not report success.");
    if (completed.seen)
    {
        state.Require(completed.sourcePath == sourcePath, L"MTP copy accounting: completion source path did not match.");
        state.Require(completed.destinationPath == destinationPath, L"MTP copy accounting: completion destination path did not match.");
    }

    std::string readBack;
    state.Require(ReadPluginFileText(io.get(), destinationPath.c_str(), readBack, state, L"MTP copy accounting read"),
                  L"MTP copy accounting: copied destination was not readable.");
    state.Require(readBack == "RedSalamander deterministic MTP fixture\r\n", L"MTP copy accounting: copied payload did not match.");

    const char* properties = nullptr;
    const HRESULT propsHr  = io->GetItemProperties(destinationPath.c_str(), &properties);
    state.Require(SUCCEEDED(propsHr) && properties != nullptr,
                  std::format(L"MTP copy accounting: GetItemProperties failed. hr=0x{:08X}", static_cast<unsigned long>(propsHr)));
    if (SUCCEEDED(propsHr) && properties)
    {
        const std::string_view props(properties);
        state.Require(extractJsonUInt(props, "copyItemCalls") == 1u,
                      L"MTP copy accounting: fake backend did not record exactly one copy call.");
        state.Require(extractJsonUInt(props, "lastCopyBytes") == 41u,
                      L"MTP copy accounting: fake backend did not record the copied byte count.");
    }

    return state.failure.empty();
});

SelfTest::RunCase(options,
                  suite,
                  L"mtp_fake_backend_move_rejects_directory_transfer_fallback",
                  [&](SelfTest::CaseState& state) noexcept
{
    CreatedFileSystemInstance created;
    const HRESULT createHr = TryCreateFakeMtpFileSystemInstance("{}", R"json({"readOnly":false})json", L"/", created);
    if (createHr == HRESULT_FROM_WIN32(ERROR_PROC_NOT_FOUND))
    {
        return state.Skip(L"MTP fake backend export is available only in debug plugin builds.");
    }
    state.Require(SUCCEEDED(createHr) && created.fileSystem,
                  std::format(L"MTP fake directory move contract: create selftest instance failed. hr=0x{:08X}", static_cast<unsigned long>(createHr)));
    if (FAILED(createHr) || ! created.fileSystem)
    {
        return false;
    }

    wil::com_ptr<IFileSystemIO> io;
    wil::com_ptr<IFileSystemDirectoryOperations> dirOps;
    state.Require(CreateFileSystemIo(created.fileSystem, io), L"MTP fake directory move contract: missing IFileSystemIO.");
    state.Require(CreateFileSystemDirectoryOperations(created.fileSystem, dirOps), L"MTP fake directory move contract: missing IFileSystemDirectoryOperations.");
    if (! io || ! dirOps)
    {
        return false;
    }

    const std::wstring guid = MakeGuidText();
    state.Require(! guid.empty(), L"MTP fake directory move contract: failed to generate unique case name.");
    if (guid.empty())
    {
        return false;
    }

    const std::wstring caseRoot      = std::format(L"/Fake Phone [devpuid:fake-device]/Internal Storage/DCIM/Camera/move-contract-{}", guid);
    const std::wstring sourceParent  = caseRoot + L"/source-parent";
    const std::wstring destParent    = caseRoot + L"/dest-parent";
    const std::wstring sourceDir     = sourceParent + L"/source-dir";
    const std::wstring destinationDir = destParent + L"/renamed-dir";
    const std::wstring destinationFile = destParent + L"/existing-file.txt";
    state.Require(SUCCEEDED(dirOps->CreateDirectory(caseRoot.c_str())), L"MTP fake directory move contract: CreateDirectory(caseRoot) failed.");
    state.Require(SUCCEEDED(dirOps->CreateDirectory(sourceParent.c_str())), L"MTP fake directory move contract: CreateDirectory(sourceParent) failed.");
    state.Require(SUCCEEDED(dirOps->CreateDirectory(destParent.c_str())), L"MTP fake directory move contract: CreateDirectory(destParent) failed.");
    state.Require(SUCCEEDED(dirOps->CreateDirectory(sourceDir.c_str())), L"MTP fake directory move contract: CreateDirectory(sourceDir) failed.");
    if (! state.failure.empty())
    {
        return false;
    }

    auto cleanup = wil::scope_exit([&]() noexcept
    { static_cast<void>(created.fileSystem->DeleteItem(caseRoot.c_str(), FILESYSTEM_FLAG_RECURSIVE, nullptr, nullptr, nullptr)); });

    constexpr std::string_view kPayload = "directory move source child";
    const std::wstring childPath        = sourceDir + L"/child.txt";
    state.Require(WritePluginFileText(io.get(), childPath.c_str(), FILESYSTEM_FLAG_NONE, kPayload, state, L"MTP fake directory move contract child write"),
                  L"MTP fake directory move contract: child write failed.");
    if (! state.failure.empty())
    {
        return false;
    }

    const HRESULT moveHr = created.fileSystem->MoveItem(sourceDir.c_str(), destinationDir.c_str(), FILESYSTEM_FLAG_NONE, nullptr, nullptr, nullptr);
    state.Require(moveHr == HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED),
                  std::format(L"MTP fake directory move contract: non-leaf-preserving directory move should be unsupported, got hr=0x{:08X}.",
                              static_cast<unsigned long>(moveHr)));

    unsigned long attrs         = 0;
    const HRESULT sourceAttrsHr = io->GetAttributes(sourceDir.c_str(), &attrs);
    state.Require(SUCCEEDED(sourceAttrsHr) && (attrs & FILE_ATTRIBUTE_DIRECTORY) != 0,
                  std::format(L"MTP fake directory move contract: source directory should remain, got hr=0x{:08X}.",
                              static_cast<unsigned long>(sourceAttrsHr)));

    const HRESULT destAttrsHr = io->GetAttributes(destinationDir.c_str(), &attrs);
    state.Require(destAttrsHr == HRESULT_FROM_WIN32(ERROR_FILE_NOT_FOUND),
                  std::format(L"MTP fake directory move contract: destination should not exist, got hr=0x{:08X}.",
                              static_cast<unsigned long>(destAttrsHr)));

    std::string readBack;
    state.Require(ReadPluginFileText(io.get(), childPath.c_str(), readBack, state, L"MTP fake directory move contract child readback"),
                  L"MTP fake directory move contract: source child readback failed.");
    state.Require(readBack == kPayload, L"MTP fake directory move contract: source child contents changed.");

    constexpr std::string_view kDestinationPayload = "directory overwrite destination must remain intact";
    state.Require(WritePluginFileText(io.get(),
                                      destinationFile.c_str(),
                                      FILESYSTEM_FLAG_NONE,
                                      kDestinationPayload,
                                      state,
                                      L"MTP fake directory overwrite destination write"),
                  L"MTP fake directory overwrite contract: destination write failed.");
    if (! state.failure.empty())
    {
        return false;
    }

    enum class DirectoryOverwriteOperation : uint8_t
    {
        Copy,
        Move,
        Rename,
    };
    constexpr std::array<DirectoryOverwriteOperation, 3> kDirectoryOverwriteOperations{{
        DirectoryOverwriteOperation::Copy,
        DirectoryOverwriteOperation::Move,
        DirectoryOverwriteOperation::Rename,
    }};
    for (const DirectoryOverwriteOperation operation : kDirectoryOverwriteOperations)
    {
        const std::wstring_view label = operation == DirectoryOverwriteOperation::Copy ? L"copy"
                                      : operation == DirectoryOverwriteOperation::Move ? L"move"
                                                                                       : L"rename";
        HRESULT operationHr = E_UNEXPECTED;
        switch (operation)
        {
        case DirectoryOverwriteOperation::Copy:
            operationHr = created.fileSystem->CopyItem(
                sourceDir.c_str(), destinationFile.c_str(), FILESYSTEM_FLAG_ALLOW_OVERWRITE, nullptr, nullptr, nullptr);
            break;
        case DirectoryOverwriteOperation::Move:
            operationHr = created.fileSystem->MoveItem(
                sourceDir.c_str(), destinationFile.c_str(), FILESYSTEM_FLAG_ALLOW_OVERWRITE, nullptr, nullptr, nullptr);
            break;
        case DirectoryOverwriteOperation::Rename:
            operationHr = created.fileSystem->RenameItem(
                sourceDir.c_str(), destinationFile.c_str(), FILESYSTEM_FLAG_ALLOW_OVERWRITE, nullptr, nullptr, nullptr);
            break;
        }

        state.Require(operationHr == HRESULT_FROM_WIN32(ERROR_ACCESS_DENIED),
                      std::format(L"MTP fake directory overwrite contract: {} should fail with ERROR_ACCESS_DENIED, got hr=0x{:08X}.",
                                  label,
                                  static_cast<unsigned long>(operationHr)));
        state.Require(SUCCEEDED(io->GetAttributes(sourceDir.c_str(), &attrs)) && (attrs & FILE_ATTRIBUTE_DIRECTORY) != 0,
                      std::format(L"MTP fake directory overwrite contract: {} removed the source directory.", label));

        readBack.clear();
        state.Require(ReadPluginFileText(io.get(),
                                         destinationFile.c_str(),
                                         readBack,
                                         state,
                                         L"MTP fake directory overwrite destination readback"),
                      std::format(L"MTP fake directory overwrite contract: {} made the destination unreadable.", label));
        state.Require(readBack == kDestinationPayload,
                      std::format(L"MTP fake directory overwrite contract: {} changed the destination contents.", label));
    }

    return state.failure.empty();
});

SelfTest::RunCase(options,
                  suite,
                  L"mtp_fake_backend_mutations_roundtrip",
                  [&](SelfTest::CaseState& state) noexcept
{
    CreatedFileSystemInstance created;
    const HRESULT createHr = TryCreateFakeMtpFileSystemInstance("{}", R"json({"readOnly":false})json", L"/", created);
    if (createHr == HRESULT_FROM_WIN32(ERROR_PROC_NOT_FOUND))
    {
        return state.Skip(L"MTP fake backend export is available only in debug plugin builds.");
    }
    state.Require(SUCCEEDED(createHr) && created.fileSystem,
                  std::format(L"MTP fake mutations: create selftest instance failed. hr=0x{:08X}", static_cast<unsigned long>(createHr)));
    if (FAILED(createHr) || ! created.fileSystem)
    {
        return false;
    }

    wil::com_ptr<IFileSystemIO> io;
    wil::com_ptr<IFileSystemDirectoryOperations> dirOps;
    state.Require(CreateFileSystemIo(created.fileSystem, io), L"MTP fake mutations: missing IFileSystemIO.");
    state.Require(CreateFileSystemDirectoryOperations(created.fileSystem, dirOps), L"MTP fake mutations: missing IFileSystemDirectoryOperations.");
    if (! io || ! dirOps)
    {
        return false;
    }

    const std::wstring guid = MakeGuidText();
    state.Require(! guid.empty(), L"MTP fake mutations: failed to generate unique case name.");
    const std::wstring caseRoot = std::format(L"/Fake Phone [devpuid:fake-device]/Internal Storage/DCIM/Camera/selftest-{}", guid);
    const HRESULT mkdirHr       = dirOps->CreateDirectory(caseRoot.c_str());
    state.Require(SUCCEEDED(mkdirHr), std::format(L"MTP fake mutations: CreateDirectory failed. hr=0x{:08X}", static_cast<unsigned long>(mkdirHr)));
    if (FAILED(mkdirHr))
    {
        return false;
    }

    auto cleanup = wil::scope_exit([&]() noexcept
    { static_cast<void>(created.fileSystem->DeleteItem(caseRoot.c_str(), FILESYSTEM_FLAG_RECURSIVE, nullptr, nullptr, nullptr)); });

    const std::wstring stagedPath = caseRoot + L"/staged.txt";
    wil::com_ptr<IFileWriter> writer;
    const HRESULT writerHr = io->CreateFileWriter(stagedPath.c_str(), FILESYSTEM_FLAG_NONE, writer.put());
    state.Require(SUCCEEDED(writerHr) && writer,
                  std::format(L"MTP fake mutations: CreateFileWriter failed. hr=0x{:08X}", static_cast<unsigned long>(writerHr)));
    if (FAILED(writerHr) || ! writer)
    {
        return false;
    }

    constexpr std::string_view kPayload = "mtp staged payload";
    const unsigned long payloadBytes    = static_cast<unsigned long>(kPayload.size());
    unsigned long written               = 0;
    const HRESULT writeHr               = writer->Write(kPayload.data(), payloadBytes, &written);
    state.Require(
        SUCCEEDED(writeHr) && written == payloadBytes,
        std::format(L"MTP fake mutations: staged write expected {} bytes, got {} hr=0x{:08X}.", payloadBytes, written, static_cast<unsigned long>(writeHr)));
    unsigned long attrs            = 0;
    const HRESULT preCommitAttrsHr = io->GetAttributes(stagedPath.c_str(), &attrs);
    state.Require(preCommitAttrsHr == HRESULT_FROM_WIN32(ERROR_FILE_NOT_FOUND),
                  std::format(L"MTP fake mutations: staged file became visible before Commit. hr=0x{:08X}", static_cast<unsigned long>(preCommitAttrsHr)));
    const HRESULT commitHr = writer->Commit();
    state.Require(SUCCEEDED(commitHr), std::format(L"MTP fake mutations: Commit failed. hr=0x{:08X}", static_cast<unsigned long>(commitHr)));
    writer.reset();
    if (! state.failure.empty())
    {
        return false;
    }

    std::string readBack;
    state.Require(ReadPluginFileText(io.get(), stagedPath.c_str(), readBack, state, L"MTP fake mutations staged read"),
                  L"MTP fake mutations: failed to read committed staged file.");
    state.Require(readBack == kPayload, L"MTP fake mutations: committed staged file contents did not match.");

    constexpr std::string_view kOverwritePayload = "mtp overwritten payload";
    state.Require(WritePluginFileText(io.get(), stagedPath.c_str(), FILESYSTEM_FLAG_ALLOW_OVERWRITE, kOverwritePayload, state, L"MTP fake mutations overwrite"),
                  L"MTP fake mutations: overwrite write failed.");
    state.Require(ReadPluginFileText(io.get(), stagedPath.c_str(), readBack, state, L"MTP fake mutations overwrite read"),
                  L"MTP fake mutations: failed to read overwritten file.");
    state.Require(readBack == kOverwritePayload, L"MTP fake mutations: overwritten file contents did not match.");

    const std::wstring copiedPath = caseRoot + L"/copied.txt";
    const HRESULT copyHr          = created.fileSystem->CopyItem(stagedPath.c_str(), copiedPath.c_str(), FILESYSTEM_FLAG_NONE, nullptr, nullptr, nullptr);
    state.Require(SUCCEEDED(copyHr), std::format(L"MTP fake mutations: CopyItem failed. hr=0x{:08X}", static_cast<unsigned long>(copyHr)));
    state.Require(ReadPluginFileText(io.get(), copiedPath.c_str(), readBack, state, L"MTP fake mutations copy read"),
                  L"MTP fake mutations: failed to read copied file.");
    state.Require(readBack == kOverwritePayload, L"MTP fake mutations: copied file contents did not match.");

    const std::wstring movedPath = caseRoot + L"/moved.txt";
    const HRESULT moveHr         = created.fileSystem->MoveItem(copiedPath.c_str(), movedPath.c_str(), FILESYSTEM_FLAG_NONE, nullptr, nullptr, nullptr);
    state.Require(SUCCEEDED(moveHr), std::format(L"MTP fake mutations: MoveItem failed. hr=0x{:08X}", static_cast<unsigned long>(moveHr)));
    state.Require(io->GetAttributes(copiedPath.c_str(), &attrs) == HRESULT_FROM_WIN32(ERROR_FILE_NOT_FOUND),
                  L"MTP fake mutations: MoveItem left the source file behind.");
    state.Require(SUCCEEDED(io->GetAttributes(movedPath.c_str(), &attrs)) && (attrs & FILE_ATTRIBUTE_DIRECTORY) == 0,
                  L"MTP fake mutations: MoveItem destination was not readable as a file.");

    const std::wstring renamedPath = caseRoot + L"/renamed.txt";
    const HRESULT renameHr         = created.fileSystem->RenameItem(movedPath.c_str(), renamedPath.c_str(), FILESYSTEM_FLAG_NONE, nullptr, nullptr, nullptr);
    state.Require(SUCCEEDED(renameHr), std::format(L"MTP fake mutations: RenameItem failed. hr=0x{:08X}", static_cast<unsigned long>(renameHr)));
    state.Require(io->GetAttributes(movedPath.c_str(), &attrs) == HRESULT_FROM_WIN32(ERROR_FILE_NOT_FOUND),
                  L"MTP fake mutations: RenameItem left the source file behind.");
    state.Require(SUCCEEDED(io->GetAttributes(renamedPath.c_str(), &attrs)) && (attrs & FILE_ATTRIBUTE_DIRECTORY) == 0,
                  L"MTP fake mutations: RenameItem destination was not readable as a file.");

    RecordingFileSystemCallback deleteCallback(1);
    const HRESULT deleteFileHr = created.fileSystem->DeleteItem(renamedPath.c_str(), FILESYSTEM_FLAG_NONE, nullptr, &deleteCallback, nullptr);
    state.Require(SUCCEEDED(deleteFileHr), std::format(L"MTP fake mutations: DeleteItem(file) failed. hr=0x{:08X}", static_cast<unsigned long>(deleteFileHr)));
    state.Require(deleteCallback.CompletedCount() == 1u,
                  std::format(L"MTP fake mutations: DeleteItem expected 1 completion callback, got {}.", deleteCallback.CompletedCount()));
    RecordedFileSystemItem completedDelete{};
    state.Require(deleteCallback.TryGetItem(0, completedDelete) && SUCCEEDED(completedDelete.status),
                  L"MTP fake mutations: DeleteItem completion callback did not report success.");
    state.Require(io->GetAttributes(renamedPath.c_str(), &attrs) == HRESULT_FROM_WIN32(ERROR_FILE_NOT_FOUND),
                  L"MTP fake mutations: DeleteItem(file) left the target behind.");

    const HRESULT deleteRootHr = created.fileSystem->DeleteItem(caseRoot.c_str(), FILESYSTEM_FLAG_RECURSIVE, nullptr, nullptr, nullptr);
    state.Require(SUCCEEDED(deleteRootHr),
                  std::format(L"MTP fake mutations: DeleteItem(directory recursive) failed. hr=0x{:08X}", static_cast<unsigned long>(deleteRootHr)));
    cleanup.release();
    state.Require(io->GetAttributes(caseRoot.c_str(), &attrs) == HRESULT_FROM_WIN32(ERROR_FILE_NOT_FOUND),
                  L"MTP fake mutations: DeleteItem(directory recursive) left the target behind.");

    return state.failure.empty();
});

SelfTest::RunCase(options,
                  suite,
                  L"mtp_fake_backend_readonly_configuration_blocks_mutations",
                  [&](SelfTest::CaseState& state) noexcept
{
    CreatedFileSystemInstance created;
    const HRESULT createHr = TryCreateFakeMtpFileSystemInstance("{}", R"json({"readOnly":true})json", L"/", created);
    if (createHr == HRESULT_FROM_WIN32(ERROR_PROC_NOT_FOUND))
    {
        return state.Skip(L"MTP fake backend export is available only in debug plugin builds.");
    }
    state.Require(SUCCEEDED(createHr) && created.fileSystem,
                  std::format(L"MTP fake read-only: create selftest instance failed. hr=0x{:08X}", static_cast<unsigned long>(createHr)));
    if (FAILED(createHr) || ! created.fileSystem)
    {
        return false;
    }

    wil::com_ptr<IFileSystemIO> io;
    wil::com_ptr<IFileSystemDirectoryOperations> dirOps;
    state.Require(CreateFileSystemIo(created.fileSystem, io), L"MTP fake read-only: missing IFileSystemIO.");
    state.Require(CreateFileSystemDirectoryOperations(created.fileSystem, dirOps), L"MTP fake read-only: missing IFileSystemDirectoryOperations.");
    if (! io || ! dirOps)
    {
        return false;
    }

    const char* capabilities = nullptr;
    const HRESULT capsHr     = created.fileSystem->GetCapabilities(&capabilities);
    state.Require(SUCCEEDED(capsHr) && capabilities != nullptr,
                  std::format(L"MTP fake read-only: GetCapabilities failed. hr=0x{:08X}", static_cast<unsigned long>(capsHr)));
    if (SUCCEEDED(capsHr) && capabilities)
    {
        const std::string_view caps(capabilities);
        state.Require(caps.find(R"json("write": false)json") != std::string_view::npos, L"MTP fake read-only: capabilities advertised write.");
        state.Require(caps.find(R"json("copy": false)json") != std::string_view::npos, L"MTP fake read-only: capabilities advertised copy.");
        state.Require(caps.find(R"json("import": { "copy": [])json") != std::string_view::npos, L"MTP fake read-only: capabilities allowed imports.");
    }

    constexpr std::wstring_view kFilePath = L"/Fake Phone [devpuid:fake-device]/Internal Storage/DCIM/Camera/photo001.txt";
    wil::com_ptr<IFileWriter> writer;
    const HRESULT writerHr = io->CreateFileWriter(kFilePath.data(), FILESYSTEM_FLAG_ALLOW_OVERWRITE, writer.put());
    state.Require(writerHr == HRESULT_FROM_WIN32(ERROR_WRITE_PROTECT),
                  std::format(L"MTP fake read-only: CreateFileWriter expected ERROR_WRITE_PROTECT, got 0x{:08X}.", static_cast<unsigned long>(writerHr)));

    const HRESULT mkdirHr = dirOps->CreateDirectory(L"/Fake Phone [devpuid:fake-device]/Internal Storage/DCIM/Camera/blocked");
    state.Require(mkdirHr == HRESULT_FROM_WIN32(ERROR_WRITE_PROTECT),
                  std::format(L"MTP fake read-only: CreateDirectory expected ERROR_WRITE_PROTECT, got 0x{:08X}.", static_cast<unsigned long>(mkdirHr)));

    const HRESULT deleteHr = created.fileSystem->DeleteItem(kFilePath.data(), FILESYSTEM_FLAG_NONE, nullptr, nullptr, nullptr);
    state.Require(deleteHr == HRESULT_FROM_WIN32(ERROR_WRITE_PROTECT),
                  std::format(L"MTP fake read-only: DeleteItem expected ERROR_WRITE_PROTECT, got 0x{:08X}.", static_cast<unsigned long>(deleteHr)));

    std::string readBack;
    state.Require(ReadPluginFileText(io.get(), kFilePath.data(), readBack, state, L"MTP fake read-only read"),
                  L"MTP fake read-only: read path failed despite read-only configuration.");
    state.Require(readBack == "RedSalamander deterministic MTP fixture\r\n", L"MTP fake read-only: readable fixture changed unexpectedly.");

    return state.failure.empty();
});

SelfTest::RunCase(options,
                  suite,
                  L"mtp_fake_backend_injection_uses_isolated_instance",
                  [&](SelfTest::CaseState& state) noexcept
{
    CreatedFileSystemInstance first;
    CreatedFileSystemInstance second;
    const HRESULT firstHr  = TryCreateFakeMtpFileSystemInstance("{}", R"json({"readOnly":false})json", L"/", first);
    const HRESULT secondHr = (firstHr == HRESULT_FROM_WIN32(ERROR_PROC_NOT_FOUND))
                                 ? firstHr
                                 : TryCreateFakeMtpFileSystemInstance("{}", R"json({"readOnly":false})json", L"/", second);
    if (firstHr == HRESULT_FROM_WIN32(ERROR_PROC_NOT_FOUND))
    {
        return state.Skip(L"MTP fake backend export is available only in debug plugin builds.");
    }
    state.Require(SUCCEEDED(firstHr) && first.fileSystem,
                  std::format(L"MTP fake isolation: first create failed. hr=0x{:08X}", static_cast<unsigned long>(firstHr)));
    state.Require(SUCCEEDED(secondHr) && second.fileSystem,
                  std::format(L"MTP fake isolation: second create failed. hr=0x{:08X}", static_cast<unsigned long>(secondHr)));
    if (FAILED(firstHr) || FAILED(secondHr) || ! first.fileSystem || ! second.fileSystem)
    {
        return false;
    }

    wil::com_ptr<IFileSystemIO> firstIo;
    wil::com_ptr<IFileSystemIO> secondIo;
    state.Require(CreateFileSystemIo(first.fileSystem, firstIo), L"MTP fake isolation: first instance missing IFileSystemIO.");
    state.Require(CreateFileSystemIo(second.fileSystem, secondIo), L"MTP fake isolation: second instance missing IFileSystemIO.");
    if (! firstIo || ! secondIo)
    {
        return false;
    }

    const std::wstring guid = MakeGuidText();
    state.Require(! guid.empty(), L"MTP fake isolation: failed to generate unique file name.");
    const std::wstring injectedPath     = std::format(L"/Fake Phone [devpuid:fake-device]/Internal Storage/DCIM/Camera/isolated-{}.txt", guid);
    constexpr std::string_view kPayload = "first instance only";
    state.Require(WritePluginFileText(firstIo.get(), injectedPath.c_str(), FILESYSTEM_FLAG_NONE, kPayload, state, L"MTP fake isolation write"),
                  L"MTP fake isolation: failed to write first instance payload.");
    const auto cleanup = wil::scope_exit([&]() noexcept
    { static_cast<void>(first.fileSystem->DeleteItem(injectedPath.c_str(), FILESYSTEM_FLAG_NONE, nullptr, nullptr, nullptr)); });

    std::string readBack;
    state.Require(ReadPluginFileText(firstIo.get(), injectedPath.c_str(), readBack, state, L"MTP fake isolation first read"),
                  L"MTP fake isolation: first instance could not read injected payload.");
    state.Require(readBack == kPayload, L"MTP fake isolation: first instance payload did not match.");

    unsigned long attrs         = 0;
    const HRESULT secondAttrsHr = secondIo->GetAttributes(injectedPath.c_str(), &attrs);
    state.Require(secondAttrsHr == HRESULT_FROM_WIN32(ERROR_FILE_NOT_FOUND),
                  std::format(L"MTP fake isolation: second instance saw first instance mutation. hr=0x{:08X}", static_cast<unsigned long>(secondAttrsHr)));

    return state.failure.empty();
});
