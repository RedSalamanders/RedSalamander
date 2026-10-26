[[nodiscard]] bool RunFileOpsCaseOnlyFileRenameStorageChecks(SelfTestState& state) noexcept
{
    const auto reject = [](std::wstring_view reason) noexcept
    {
        AppendLog(std::format(L"Case-only file rename storage: {}", reason));
        return false;
    };
    auto* folder     = TryGetFolderWindow(state.mainWindow);
    auto* operations = TryGetFileOps(folder);
    if (! folder || ! operations || ! CancelOutstandingSelfTestOperations(operations))
        return reject(L"owner or initial worker quiet point unavailable");
    RedSalamander::TestSupport::ScopedWindowActivationBlocker blocker;
    if (! RedSalamander::TestSupport::ScopedWindowActivationBlocker::IsActiveForCurrentThread() && ! blocker.Start())
        return reject(L"no-activation guard unavailable");
    const HWND foreground    = GetForegroundWindow();
    const HWND originalPopup = operations->GetPopupHwndForSelfTest();
    const auto fileSystem    = folder->GetFileSystem(FolderWindow::Pane::Left);
    wil::com_ptr<IFileSystemIO> io;
    if (! fileSystem || folder->GetFileSystemPluginId(FolderWindow::Pane::Left) != kPluginIdLocal || ! fileSystem.try_query_to(io.put()))
        return reject(L"actual Local pane provider unavailable");

    const auto root                = state.tempRoot / L"case-only-file-rename-storage";
    const auto hostSource          = root / L"host" / L"MixedCase_r\u00E9vision_\u6771\u4EAC.BIN";
    const auto hostDestination     = hostSource.parent_path() / L"mixedcase_r\u00E9vision_\u6771\u4EAC.bin";
    const auto providerSource      = root / L"provider" / L"Rollback_r\u00E9vision_\u6771\u4EAC.BIN";
    const auto providerDestination = providerSource.parent_path() / L"rollback_r\u00E9vision_\u6771\u4EAC.bin";
    const auto hostReference       = root / L"references" / L"host.bin";
    const auto providerReference   = root / L"references" / L"provider.bin";
    if (! RecreateEmptyDirectory(root))
        return reject(L"case-owned fixture root unavailable");
    uint64_t taskId       = 0u;
    bool cleanupQuiet     = true;
    const auto removeRoot = wil::scope_exit([&]() noexcept
    {
        if (cleanupQuiet)
            static_cast<void>(SelfTest::RemoveAll(root));
        else
            AppendLog(L"Case-only rename cleanup: worker retirement failed; owned fixture retained.");
    });
    if (! WriteFilledTestFile(hostSource, 4099u, 0xA7u) || ! WriteFilledTestFile(hostReference, 4099u, 0xA7u) ||
        ! WriteFilledTestFile(providerSource, 2053u, 0x3Cu) || ! WriteFilledTestFile(providerReference, 2053u, 0x3Cu))
        return reject(L"case-owned distinct payloads/references unavailable");
    // This fixed local fixture needs the native FILE_ID_INFO as independent storage
    // evidence, not a general cross-provider identity utility. Open the exact entry
    // no-follow and reject links/directories before querying its file ID.
    const auto readId = [](const std::filesystem::path& path, FILE_ID_INFO& id) noexcept
    {
        const auto extended = MakeExtendedPathForSelfTest(path);
        wil::unique_hfile file(CreateFileW(extended.c_str(),
                                           FILE_READ_ATTRIBUTES,
                                           FILE_SHARE_READ | FILE_SHARE_WRITE | FILE_SHARE_DELETE,
                                           nullptr,
                                           OPEN_EXISTING,
                                           FILE_FLAG_OPEN_REPARSE_POINT,
                                           nullptr));
        FILE_ATTRIBUTE_TAG_INFO attributes{};
        return file && GetFileInformationByHandleEx(file.get(), FileAttributeTagInfo, &attributes, sizeof(attributes)) &&
               (attributes.FileAttributes & (FILE_ATTRIBUTE_DIRECTORY | FILE_ATTRIBUTE_REPARSE_POINT)) == 0u &&
               GetFileInformationByHandleEx(file.get(), FileIdInfo, &id, sizeof(id));
    };
    const auto sameId = [](const FILE_ID_INFO& left, const FILE_ID_INFO& right) noexcept
    {
        return left.VolumeSerialNumber == right.VolumeSerialNumber &&
               std::memcmp(left.FileId.Identifier, right.FileId.Identifier, sizeof(left.FileId.Identifier)) == 0;
    };
    const auto oneExactLeaf = [](const std::filesystem::path& parent, std::wstring_view expected) noexcept
    {
        const DWORD attributes = GetFileAttributesW(MakeExtendedPathForSelfTest(parent).c_str());
        if (attributes == INVALID_FILE_ATTRIBUTES || (attributes & FILE_ATTRIBUTE_DIRECTORY) == 0u || (attributes & FILE_ATTRIBUTE_REPARSE_POINT) != 0u)
            return false;
        std::error_code ec;
        std::filesystem::directory_iterator entry(parent, std::filesystem::directory_options::none, ec);
        size_t count = 0u;
        for (; ! ec && entry != std::filesystem::directory_iterator{}; entry.increment(ec))
        {
            const auto path             = entry->path();
            const DWORD childAttributes = GetFileAttributesW(MakeExtendedPathForSelfTest(path).c_str());
            if (++count > 1u || path.filename().native() != expected || childAttributes == INVALID_FILE_ATTRIBUTES ||
                (childAttributes & (FILE_ATTRIBUTE_DIRECTORY | FILE_ATTRIBUTE_REPARSE_POINT)) != 0u)
                return false;
        }
        // Enumerated exact casing is authoritative: exists(old spelling) also succeeds
        // after case-only rename on a case-insensitive volume and proves no name change.
        return ! ec && count == 1u;
    };
    const auto digest = [](const std::string& bytes) noexcept
    {
        Common::Crypto::Blake3Hasher hasher;
        hasher.Update(std::as_bytes(std::span<const char>(bytes.data(), bytes.size())));
        return hasher.Finalize();
    };
    FILE_ID_INFO hostId{};
    FILE_ID_INFO providerId{};
    std::string hostBytes;
    std::string providerBytes;
    if (! readId(hostSource, hostId) || ! readId(providerSource, providerId) || ! ReadFileTextFsIo(io, hostSource, hostBytes) ||
        ! ReadFileTextFsIo(io, providerSource, providerBytes) || hostBytes.size() != 4099u || providerBytes.size() != 2053u ||
        ! FilesEqualBytes(hostSource, hostReference) || ! FilesEqualBytes(providerSource, providerReference) ||
        ! oneExactLeaf(hostSource.parent_path(), hostSource.filename().native()) ||
        ! oneExactLeaf(providerSource.parent_path(), providerSource.filename().native()))
        return reject(L"original file IDs, exact names or byte references unavailable");
    const auto hostDigest          = digest(hostBytes);
    const auto providerDigest      = digest(providerBytes);
    const bool autoDismiss         = operations->GetAutoDismissSuccess();
    const bool autoAccept          = HostGetAutoAcceptPrompts();
    const auto promptOverride      = HostGetTestPromptResultOverride();
    const auto restorePresentation = [&]() noexcept
    {
        HostSetAutoAcceptPrompts(autoAccept);
        if (promptOverride == HOST_PROMPT_RESULT_NONE)
            HostClearTestPromptResultOverride();
        else
            HostSetTestPromptResultOverride(promptOverride);
        operations->SetAutoDismissSuccess(autoDismiss);
        return operations->DebugFlushPendingSettingsSaveForSelfTest(5'000u);
    };
    auto restorePresentationGuard = wil::scope_exit([&]() noexcept { static_cast<void>(restorePresentation()); });
    operations->SetAutoDismissSuccess(false);
    HostSetAutoAcceptPrompts(false);
    HostClearTestPromptResultOverride();
    struct Observation final
    {
        uint64_t taskId    = 0u;
        unsigned int count = 0u;
        HRESULT result     = E_PENDING;
    };
    auto observation            = std::make_shared<Observation>();
    const uint64_t subscription = folder->AddFileOperationCompletedCallback([observation](const FolderWindow::FileOperationCompletedEvent& event) noexcept
    {
        if (event.taskId == observation->taskId)
        {
            observation->result = event.hr;
            ++observation->count;
        }
    });
    const auto drain            = wil::scope_exit([&]() noexcept
    {
        cleanupQuiet = CancelOutstandingSelfTestOperations(operations);
        if (cleanupQuiet && taskId != 0u)
        {
            const auto retired = RedSalamander::TestSupport::PumpMessagesUntil([&]() noexcept {
                return ! operations->FindTask(taskId);
            }, {.timeout = SelfTest::Scale(std::chrono::seconds(30)), .operationName = L"D8 failure cleanup worker retirement"});
            cleanupQuiet       = retired.conditionMet;
        }
        folder->RemoveFileOperationCompletedCallback(subscription);
        if (cleanupQuiet && taskId != 0u)
        {
            operations->DismissCompletedTask(taskId);
            operations->DebugRemoveDiagnosticsForTask(taskId);
            state.completedTasks.erase(taskId);
        }
        if (cleanupQuiet && ! originalPopup && operations->GetPopupHwndForSelfTest())
            operations->DebugResetPopupForSelfTest();
    });
    const auto promptsBefore    = HostGetTestPromptRequestCount();
    std::filesystem::path admittedParent;
    const HRESULT admitted =
        operations->AdmitInlineRename(FolderWindow::Pane::Left, fileSystem, hostSource, hostDestination.filename().native(), &taskId, &admittedParent);
    if (FAILED(admitted) || taskId == 0u || admittedParent != hostSource.parent_path())
        return reject(L"actual pane-bound InlineRename admission failed");
    observation->taskId = taskId;
    cleanupQuiet        = false;
    const auto quiet    = RedSalamander::TestSupport::PumpMessagesUntil([&]() noexcept {
        return observation->count == 1u && ! operations->FindTask(taskId) && ! operations->HasActiveOperations();
    }, {.timeout = SelfTest::Scale(std::chrono::seconds(8)), .operationName = L"D8 real host case-only rename post-join quiet point"});
    if (! quiet.conditionMet || observation->result != S_OK || GetForegroundWindow() != foreground || HostGetTestPromptRequestCount() != promptsBefore)
        return reject(L"real host case-only rename failed, duplicated completion, prompted or activated");
    cleanupQuiet = true;
    FILE_ID_INFO hostAfterId{};
    std::string hostAfterBytes;
    const auto completion = state.completedTasks.find(taskId);
    if (! oneExactLeaf(hostSource.parent_path(), hostDestination.filename().native()) || ! readId(hostDestination, hostAfterId) ||
        ! sameId(hostId, hostAfterId) || ! ReadFileTextFsIo(io, hostDestination, hostAfterBytes) || hostAfterBytes != hostBytes ||
        digest(hostAfterBytes) != hostDigest || ! FilesEqualBytes(hostDestination, hostReference) || completion == state.completedTasks.end() ||
        completion->second.hr != S_OK || completion->second.sourceItemResults.size() != 1u || ! completion->second.sourceItemResults.front().has_value())
        return reject(L"host final exact casing, file ID, bytes/BLAKE3, no-temp inventory or dense result failed");
    const auto& result = completion->second.sourceItemResults.front().value();
    if (result.strategy != FileOperations::OperationStrategy::Native || result.completion != FileOperations::ItemCompletion::Completed ||
        result.status != S_OK || result.publication != FileOperations::PublicationState::Published ||
        result.sourceDisposition != FileOperations::SourceDisposition::Removed || result.finalSourcePath != hostSource.native() ||
        result.finalDestinationPath != hostDestination.native())
        return reject(L"actual host typed result did not retain Completed/Published/Removed rename truth");
    AppendLog(std::format(L"D8 host task={} AdmitInlineRename/RenameIfUnchanged: exact leaf='{}', volume={:016X}, fileId={}, bytes={}, BLAKE3={}; "
                          L"identity/exact bytes retained, no .rs_case_tmp_ siblings, one S_OK callback after retirement, foreground preserved",
                          taskId,
                          hostDestination.filename().native(),
                          hostAfterId.VolumeSerialNumber,
                          Utf16FromUtf8ForFileOpsSelfTest(Common::Crypto::EncodeHexDigest(std::as_bytes(std::span{hostAfterId.FileId.Identifier}))),
                          hostAfterBytes.size(),
                          Utf16FromUtf8ForFileOpsSelfTest(Common::Crypto::EncodeHexDigest(hostDigest))));

    try
    {
        constexpr std::wstring_view failVariable     = L"REDSALAMANDER_FILEOPS_CASE_RENAME_SECOND_LEG_FAIL_PATH";
        constexpr std::wstring_view firstLegVariable = L"REDSALAMANDER_FILEOPS_CASE_RENAME_FIRST_LEG_PATH";
        constexpr std::wstring_view revertedVariable = L"REDSALAMANDER_FILEOPS_CASE_RENAME_REVERTED";
        if (! GetEnvVarTrimmed(failVariable).empty() || ! GetEnvVarTrimmed(firstLegVariable).empty() || ! GetEnvVarTrimmed(revertedVariable).empty())
            return reject(L"another provider second-leg fixture is configured");
        const auto extendedSource = MakeExtendedPathForSelfTest(providerSource);
        RedSalamander::TestSupport::ScopedEnvironmentVariable failScope(failVariable, extendedSource);
        RedSalamander::TestSupport::ScopedEnvironmentVariable firstLegScope(firstLegVariable, std::nullopt);
        RedSalamander::TestSupport::ScopedEnvironmentVariable revertedScope(revertedVariable, std::nullopt);
        // This distinct synchronous public-provider route retains the existing temp
        // fallback and ordinary revert. It is explicitly not a host InlineRename task.
        const HRESULT providerResult =
            fileSystem->RenameItem(providerSource.c_str(), providerDestination.c_str(), FILESYSTEM_FLAG_NONE, nullptr, nullptr, nullptr);
        const auto temporaryName = GetEnvVarTrimmed(firstLegVariable);
        const std::filesystem::path temporaryPath(temporaryName);
        FILE_ID_INFO providerAfterId{};
        std::string providerAfterBytes;
        if (providerResult != HRESULT_FROM_WIN32(ERROR_WRITE_FAULT) || ! GetEnvVarTrimmed(failVariable).empty() || temporaryName.empty() ||
            temporaryPath.parent_path().native() != MakeExtendedPathForSelfTest(providerSource.parent_path()) ||
            ! temporaryPath.filename().native().starts_with(L".rs_case_tmp_") || GetEnvVarTrimmed(revertedVariable) != L"1" ||
            ! oneExactLeaf(providerSource.parent_path(), providerSource.filename().native()) || ! readId(providerSource, providerAfterId) ||
            ! sameId(providerId, providerAfterId) || ! ReadFileTextFsIo(io, providerSource, providerAfterBytes) || providerAfterBytes != providerBytes ||
            digest(providerAfterBytes) != providerDigest || ! FilesEqualBytes(providerSource, providerReference) || GetForegroundWindow() != foreground ||
            operations->HasActiveOperations())
            return reject(L"provider real-first-hop/controlled-second-leg/ordinary-rollback exact name/ID/bytes/no-temp proof failed");
        AppendLog(
            std::format(L"D8 provider RenameItem/RenameCaseOnlyWithTemp: actual first-hop temporary='{}'; leg-two injected ERROR_WRITE_FAULT; "
                        L"ordinary MoveFileExW rollback succeeded, exact restored leaf='{}', volume={:016X}, fileId={}, bytes={}, BLAKE3={}; no temp siblings",
                        temporaryName,
                        providerSource.filename().native(),
                        providerAfterId.VolumeSerialNumber,
                        Utf16FromUtf8ForFileOpsSelfTest(Common::Crypto::EncodeHexDigest(std::as_bytes(std::span{providerAfterId.FileId.Identifier}))),
                        providerAfterBytes.size(),
                        Utf16FromUtf8ForFileOpsSelfTest(Common::Crypto::EncodeHexDigest(providerDigest))));
    }
    catch (const std::system_error& error)
    {
        // Canonical environment scopes throw on Win32 setup failure; restore them
        // before reporting failure at this noexcept selftest boundary.
        Debug::Error(L"D8 environment setup failed (error={}).", error.code().value());
        return reject(L"canonical environment scope setup failed");
    }
    if (! restorePresentation())
        return reject(L"restored presentation settings did not reach their quiet point");
    restorePresentationGuard.release();
    return GetForegroundWindow() == foreground;
}
