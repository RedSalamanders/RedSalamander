// Included by the ENABLE_TESTS FileOperations coordinator; never a separate TU.
[[nodiscard]] bool RunFileOpsAclDeniedChildCopyChecks(SelfTestState& state) noexcept
{
    using State       = FolderWindow::FileOperationState;
    using Task        = State::Task;
    const auto reject = [](std::wstring_view reason) noexcept
    {
        AppendLog(L"ACL denied child Copy:");
        AppendLog(reason);
        return false;
    };
    try
    {
        auto* folder     = TryGetFolderWindow(state.mainWindow);
        auto* operations = TryGetFileOps(folder);
        if (! folder || ! operations || ! CancelOutstandingSelfTestOperations(operations))
            return reject(L"owner or initial worker quiet point unavailable");
        RedSalamander::TestSupport::ScopedWindowActivationBlocker blocker;
        if (! RedSalamander::TestSupport::ScopedWindowActivationBlocker::IsActiveForCurrentThread() && ! blocker.Start())
            return reject(L"no-activation guard unavailable");
        const HWND foreground            = GetForegroundWindow();
        const HWND originalPopup         = operations->GetPopupHwndForSelfTest();
        const auto sourceFileSystem      = folder->GetFileSystem(FolderWindow::Pane::Left);
        const auto destinationFileSystem = folder->GetFileSystem(FolderWindow::Pane::Right);
        wil::com_ptr<IFileSystemIO> io;
        wil::com_ptr<IInformations> information;
        if (! sourceFileSystem || ! destinationFileSystem || folder->GetFileSystemPluginId(FolderWindow::Pane::Left) != kPluginIdLocal ||
            folder->GetFileSystemPluginId(FolderWindow::Pane::Right) != kPluginIdLocal || ! sourceFileSystem.try_query_to(io.put()) ||
            ! sourceFileSystem.try_query_to(information.put()))
            return reject(L"actual Local pane interfaces unavailable");

        const auto root      = state.tempRoot / L"acl-denied-child-copy";
        const auto source    = root / L"source_r\u00E9vision_\u6771\u4EAC";
        const auto denied    = source / L"denied-child";
        const auto reference = root / L"reference";
        const auto target    = root / L"destination";
        const auto published = target / source.filename();
        // Refuse an old retained case; it may belong to a worker whose quiet point
        // failed. Never recreate/delete its tree before ownership has been reviewed.
        std::error_code existsError;
        const bool rootExists = std::filesystem::exists(root, existsError);
        if (rootExists || existsError)
        {
            state.retainTempRootOnFailure = true;
            return reject(L"case root exists or cannot be inspected; retained data requires review");
        }
        if (! SelfTest::EnsureDirectory(root))
            return reject(L"case root exists or cannot be created; retained data requires review");
        bool cleanable        = true;
        const auto removeRoot = wil::scope_exit([&]() noexcept
        {
            if (cleanable)
                static_cast<void>(SelfTest::RemoveAll(root));
            else
            {
                state.retainTempRootOnFailure = true;
                AppendLog(L"ACL Copy cleanup retained case: worker/security/result proof incomplete.");
                AppendLog(root.native());
            }
        });
        if (! WriteFilledTestFile(root / L"case-owned.marker", 47u, 0xEEu) || ! WriteFilledTestFile(denied / L"secret.bin", 1307u, 0xD3u) ||
            ! WriteFilledTestFile(source / L"accessible.bin", 4099u, 0xA7u) || ! WriteFilledTestFile(reference / L"secret.bin", 1307u, 0xD3u) ||
            ! WriteFilledTestFile(reference / L"accessible.bin", 4099u, 0xA7u) || ! SelfTest::EnsureDirectory(target))
            return reject(L"small distinct source/reference fixture unavailable");
        // From the first storage proof onward, a refusal preserves this case for
        // review. Only the complete quiet/restoration/content proof clears it.
        cleanable = false;

        struct Entry final
        {
            std::wstring relative;
            bool directory = false;
            std::string bytes;
            Common::Crypto::Blake3Digest digest{};
            bool operator==(const Entry&) const = default;
        };
        // Closed finite inventory: enumerate only these exact directories. Check
        // every observed object for reparse attributes before IO; never follow a
        // discovered directory, even when its name resembles a permitted child.
        const auto capture = [&](bool after, std::vector<Entry>& entries)
        {
            entries.clear();
            std::vector<std::filesystem::path> directories{root, source, denied, reference, target};
            if (after)
            {
                directories.push_back(published);
                directories.push_back(published / denied.filename());
            }
            size_t totalBytes = 0u;
            for (const auto& directory : directories)
            {
                const DWORD attributes = GetFileAttributesW(directory.c_str());
                if (attributes == INVALID_FILE_ATTRIBUTES || (attributes & FILE_ATTRIBUTE_REPARSE_POINT) != 0u || (attributes & FILE_ATTRIBUTE_DIRECTORY) == 0u)
                    return false;
                std::error_code ec;
                std::filesystem::directory_iterator item(directory, std::filesystem::directory_options::none, ec);
                if (ec)
                    return false;
                for (; item != std::filesystem::directory_iterator{}; item.increment(ec))
                {
                    if (ec || entries.size() >= 12u)
                        return false;
                    const auto path            = item->path();
                    const DWORD itemAttributes = GetFileAttributesW(path.c_str());
                    if (itemAttributes == INVALID_FILE_ATTRIBUTES || (itemAttributes & FILE_ATTRIBUTE_REPARSE_POINT) != 0u)
                        return false;
                    const auto itemStatus = item->symlink_status(ec);
                    if (ec || (! std::filesystem::is_directory(itemStatus) && ! std::filesystem::is_regular_file(itemStatus)))
                        return false;
                    Entry entry;
                    entry.relative  = path.lexically_relative(root).native();
                    entry.directory = (itemAttributes & FILE_ATTRIBUTE_DIRECTORY) != 0u;
                    if (! entry.directory)
                    {
                        WIN32_FILE_ATTRIBUTE_DATA data{};
                        if (! GetFileAttributesExW(path.c_str(), GetFileExInfoStandard, &data) || data.nFileSizeHigh != 0u || data.nFileSizeLow > 4099u ||
                            totalBytes + data.nFileSizeLow > 24'000u || ! ReadFileTextFsIo(io, path, entry.bytes) || entry.bytes.size() != data.nFileSizeLow)
                            return false;
                        totalBytes += entry.bytes.size();
                        Common::Crypto::Blake3Hasher hasher;
                        hasher.Update(std::as_bytes(std::span<const char>(entry.bytes.data(), entry.bytes.size())));
                        entry.digest = hasher.Finalize();
                    }
                    entries.push_back(std::move(entry));
                }
                if (ec)
                    return false;
            }
            std::ranges::sort(entries, {}, &Entry::relative);
            return true;
        };
        std::vector<Entry> before;
        std::vector<std::wstring> expectedNames{L"case-owned.marker",
                                                source.filename().native(),
                                                denied.lexically_relative(root).native(),
                                                (denied / L"secret.bin").lexically_relative(root).native(),
                                                (source / L"accessible.bin").lexically_relative(root).native(),
                                                L"reference",
                                                L"reference\\secret.bin",
                                                L"reference\\accessible.bin",
                                                L"destination"};
        std::ranges::sort(expectedNames);
        if (! capture(false, before) || before.size() != expectedNames.size())
            return reject(L"before-operation closed inventory unavailable");
        for (size_t index = 0u; index < before.size(); ++index)
        {
            if (before[index].relative != expectedNames[index])
                return reject(L"before-operation inventory differs from exact owned paths");
        }
        struct Seed final
        {
            std::wstring relative;
            size_t size;
            unsigned char value;
        };
        const std::array<Seed, 4> seeds{{
            {(denied / L"secret.bin").lexically_relative(root).native(), 1307u, 0xD3u},
            {(source / L"accessible.bin").lexically_relative(root).native(), 4099u, 0xA7u},
            {L"reference\\secret.bin", 1307u, 0xD3u},
            {L"reference\\accessible.bin", 4099u, 0xA7u},
        }};
        for (const auto& [relative, size, value] : seeds)
        {
            const auto entry = std::ranges::find(before, relative, &Entry::relative);
            if (entry == before.end() || entry->directory || entry->bytes.size() != size ||
                ! std::ranges::all_of(entry->bytes, [value](char byte) noexcept { return static_cast<unsigned char>(byte) == value; }))
                return reject(L"before-operation distinct exact payload proof failed");
        }
        for (const auto& entry : before)
        {
            AppendLog(std::format(L"ACL Copy closed inventory before: {} kind={} bytes={} BLAKE3={}",
                                  entry.relative,
                                  entry.directory ? L"ordinary-directory" : L"ordinary-file",
                                  entry.bytes.size(),
                                  entry.directory ? std::wstring{} : Utf16FromUtf8ForFileOpsSelfTest(Common::Crypto::EncodeHexDigest(entry.digest))));
        }

        // Retain authority before denial. Restore this exact directory, not a
        // path reopened after the worker. No token privilege change or ACL grant.
        wil::unique_handle securityHandle(CreateFileW(denied.c_str(),
                                                      READ_CONTROL | WRITE_DAC | FILE_READ_ATTRIBUTES,
                                                      FILE_SHARE_READ | FILE_SHARE_WRITE | FILE_SHARE_DELETE,
                                                      nullptr,
                                                      OPEN_EXISTING,
                                                      FILE_FLAG_OPEN_REPARSE_POINT | FILE_FLAG_BACKUP_SEMANTICS,
                                                      nullptr));
        FILE_ATTRIBUTE_TAG_INFO attributeTag{};
        if (! securityHandle || ! GetFileInformationByHandleEx(securityHandle.get(), FileAttributeTagInfo, &attributeTag, sizeof(attributeTag)) ||
            (attributeTag.FileAttributes & FILE_ATTRIBUTE_DIRECTORY) == 0u || (attributeTag.FileAttributes & FILE_ATTRIBUTE_REPARSE_POINT) != 0u)
            return reject(L"exact ordinary denied-directory security authority unavailable");
        constexpr SECURITY_INFORMATION securityParts = OWNER_SECURITY_INFORMATION | GROUP_SECURITY_INFORMATION | DACL_SECURITY_INFORMATION;
        PSECURITY_DESCRIPTOR originalRaw             = nullptr;
        PACL originalDacl                            = nullptr;
        const DWORD originalError =
            GetSecurityInfo(securityHandle.get(), SE_FILE_OBJECT, securityParts, nullptr, nullptr, &originalDacl, nullptr, &originalRaw);
        wil::unique_hlocal_ptr<void> originalSecurity(originalRaw);
        SECURITY_DESCRIPTOR_CONTROL originalControl{};
        DWORD originalRevision = 0u;
        if (originalError != ERROR_SUCCESS || ! originalSecurity || ! originalDacl || ! IsValidAcl(originalDacl) ||
            ! IsValidSecurityDescriptor(originalSecurity.get()) || ! GetSecurityDescriptorControl(originalSecurity.get(), &originalControl, &originalRevision))
            return reject(L"original owner/group/DACL security descriptor unavailable; null DACL is refused");
        // Establish the new private fixture in the inheritance model used by SetSecurityInfo
        // before taking the immutable pre-denial snapshot. Otherwise the first deny/restore
        // converts SE_DACL_AUTO_INHERITED and an exact restore can never match the old model.
        // Reapply only its existing DACL and protection; no permission or privilege is granted.
        const SECURITY_INFORMATION initializeParts =
            DACL_SECURITY_INFORMATION |
            ((originalControl & SE_DACL_PROTECTED) != 0u ? PROTECTED_DACL_SECURITY_INFORMATION : UNPROTECTED_DACL_SECURITY_INFORMATION);
        const DWORD initializeError = SetSecurityInfo(securityHandle.get(), SE_FILE_OBJECT, initializeParts, nullptr, nullptr, originalDacl, nullptr);
        if (initializeError != ERROR_SUCCESS)
            return reject(std::format(L"private fixture DACL initialization failed ({}).", initializeError));
        originalSecurity.reset();
        originalRaw  = nullptr;
        originalDacl = nullptr;
        const DWORD snapshotError =
            GetSecurityInfo(securityHandle.get(), SE_FILE_OBJECT, securityParts, nullptr, nullptr, &originalDacl, nullptr, &originalRaw);
        originalSecurity.reset(originalRaw);
        if (snapshotError != ERROR_SUCCESS || ! originalSecurity || ! originalDacl || ! IsValidAcl(originalDacl) ||
            ! IsValidSecurityDescriptor(originalSecurity.get()) || ! GetSecurityDescriptorControl(originalSecurity.get(), &originalControl, &originalRevision))
            return reject(L"initialized pre-denial security snapshot unavailable");
        bool securityChanged       = false;
        const auto restoreSecurity = [&]() noexcept
        {
            if (! securityChanged)
                return true;
            const SECURITY_INFORMATION restoreParts =
                DACL_SECURITY_INFORMATION |
                ((originalControl & SE_DACL_PROTECTED) != 0u ? PROTECTED_DACL_SECURITY_INFORMATION : UNPROTECTED_DACL_SECURITY_INFORMATION);
            const DWORD restoreError = SetSecurityInfo(securityHandle.get(), SE_FILE_OBJECT, restoreParts, nullptr, nullptr, originalDacl, nullptr);
            if (restoreError != ERROR_SUCCESS)
            {
                AppendLog(std::format(L"ACL Copy restore: SetSecurityInfo failed ({}).", restoreError));
                return false;
            }
            PSECURITY_DESCRIPTOR restoredRaw = nullptr;
            PACL restoredDacl                = nullptr;
            const DWORD queryError =
                GetSecurityInfo(securityHandle.get(), SE_FILE_OBJECT, securityParts, nullptr, nullptr, &restoredDacl, nullptr, &restoredRaw);
            wil::unique_hlocal_ptr<void> restoredSecurity(restoredRaw);
            SECURITY_DESCRIPTOR_CONTROL restoredControl{};
            DWORD restoredRevision = 0u;
            if (queryError != ERROR_SUCCESS || ! restoredSecurity || ! restoredDacl || ! IsValidAcl(restoredDacl) ||
                ! IsValidSecurityDescriptor(restoredSecurity.get()) ||
                ! GetSecurityDescriptorControl(restoredSecurity.get(), &restoredControl, &restoredRevision) || restoredControl != originalControl ||
                restoredRevision != originalRevision || originalDacl->AclSize != restoredDacl->AclSize ||
                std::memcmp(originalDacl, restoredDacl, originalDacl->AclSize) != 0 ||
                GetSecurityDescriptorLength(originalSecurity.get()) != GetSecurityDescriptorLength(restoredSecurity.get()) ||
                std::memcmp(originalSecurity.get(), restoredSecurity.get(), GetSecurityDescriptorLength(originalSecurity.get())) != 0)
            {
                AppendLog(std::format(
                    L"ACL Copy restore: query={} control=0x{:04X}/0x{:04X} revision={}/{} DACL bytes={}/{} SD bytes={}/{}.",
                    queryError,
                    static_cast<unsigned int>(originalControl),
                    static_cast<unsigned int>(restoredControl),
                    originalRevision,
                    restoredRevision,
                    originalDacl->AclSize,
                    restoredDacl ? restoredDacl->AclSize : 0u,
                    GetSecurityDescriptorLength(originalSecurity.get()),
                    restoredSecurity && IsValidSecurityDescriptor(restoredSecurity.get()) ? GetSecurityDescriptorLength(restoredSecurity.get()) : 0u));
                return false;
            }
            securityChanged = false;
            return true;
        };
        const auto securityGuard = wil::scope_exit([&]() noexcept
        {
            if (! restoreSecurity())
            {
                cleanable = false;
                AppendLog(L"ACL Copy cleanup FAILED to restore exact original owner/group/DACL descriptor/control; case retained.");
            }
        });

        std::string configuration;
        if (! BackupPluginConfiguration(information.get(), configuration))
            return reject(L"actual source pane configuration backup unavailable");
        bool configurationChanged       = false;
        const auto restoreConfiguration = [&]() noexcept
        {
            if (! configurationChanged)
                return true;
            const bool restored  = SUCCEEDED(information->SetConfiguration(configuration.c_str()));
            configurationChanged = ! restored;
            return restored;
        };
        const auto configurationGuard = wil::scope_exit([&]() noexcept
        {
            if (! restoreConfiguration())
                AppendLog(L"ACL Copy cleanup FAILED to restore source provider configuration.");
        });
        configurationChanged          = true;
        if (FAILED(information->SetConfiguration(R"json({"concurrencyMode":"manual","copyMoveMaxConcurrency":1})json")))
            return reject(L"serial Local configuration refused");
        const auto sourceRoute = FileSystemRouteContract::Query(sourceFileSystem.get(), source.native(), FILESYSTEM_COPY, kPluginIdLocal);
        if (sourceRoute.state != FileSystemRouteContract::QueryState::Available || sourceRoute.snapshot.copyMoveMaxConcurrency != 1u)
            return reject(L"actual source pane route did not establish serial Copy");

        const bool autoDismiss     = operations->GetAutoDismissSuccess();
        const bool queueNew        = operations->GetQueueNewTasks();
        const bool autoAccept      = HostGetAutoAcceptPrompts();
        const bool verifyAfterCopy = EnsureFileOperationsSettingsForSelfTest().verifyAfterCopy;
        const auto overrideResult  = HostGetTestPromptResultOverride();
        const auto restoreSettings = [&]() noexcept
        {
            EnsureFileOperationsSettingsForSelfTest().verifyAfterCopy = verifyAfterCopy;
            HostSetAutoAcceptPrompts(autoAccept);
            if (overrideResult == HOST_PROMPT_RESULT_NONE)
                HostClearTestPromptResultOverride();
            else
                HostSetTestPromptResultOverride(overrideResult);
            operations->SetAutoDismissSuccess(autoDismiss);
            operations->SetQueueNewTasks(queueNew);
            return operations->DebugFlushPendingSettingsSaveForSelfTest(5'000u);
        };
        const auto settingsGuard = wil::scope_exit([&]() noexcept { static_cast<void>(restoreSettings()); });
        HostSetAutoAcceptPrompts(false);
        HostClearTestPromptResultOverride();
        operations->SetAutoDismissSuccess(false);
        operations->SetQueueNewTasks(false);
        // Select ordinary same-provider Local Copy. Requested verification or an
        // explicit destinationFileSystem selects the host bridge instead.
        EnsureFileOperationsSettingsForSelfTest().verifyAfterCopy = false;

        struct Observation final
        {
            // Completion dispatch and the test predicates run on this same owner
            // thread; shared ownership is for callback lifetime, not worker access.
            uint64_t taskId    = 0u;
            unsigned int count = 0u;
            HRESULT hr         = E_PENDING;
        };
        const auto observation      = std::make_shared<Observation>();
        const uint64_t subscription = folder->AddFileOperationCompletedCallback([observation](const FolderWindow::FileOperationCompletedEvent& event) noexcept
        {
            if (event.taskId == observation->taskId)
            {
                observation->hr = event.hr;
                ++observation->count;
            }
        });
        const auto unsubscribe      = wil::scope_exit([&]() noexcept { folder->RemoveFileOperationCompletedCallback(subscription); });
        const auto retire           = [&]() noexcept
        {
            if (! CancelOutstandingSelfTestOperations(operations))
                return false;
            if (observation->taskId == 0u)
                return true;
            const auto quiet = RedSalamander::TestSupport::PumpMessagesUntil([&]() noexcept {
                return ! operations->HasActiveOperations() && ! operations->FindTask(observation->taskId);
            }, {.timeout = SelfTest::Scale(std::chrono::seconds(8)), .operationName = L"ACL Copy exact worker retirement"});
            return quiet.conditionMet;
        };
        const auto drain = wil::scope_exit([&]() noexcept
        {
            // Shared callback storage survives unsubscribe if owner dispatch copied
            // the registration. Failure cleanup never recursively deletes live data.
            const bool quiet = retire();
            cleanable        = quiet && cleanable;
            if (! quiet)
            {
                // Terminal failure-only recovery: this fixture admitted from an
                // idle state. Join task/reaper workers before DACL, configuration,
                // prompt policy or callback guards can unwind over live work.
                AppendLog(L"ACL Copy bounded retirement failed; hard shutdown joins owned workers before restoring DACL/configuration.");
                operations->Shutdown();
            }
            if (! restoreSecurity())
                cleanable = false;
            if (quiet && cleanable && observation->taskId != 0u)
            {
                operations->DismissCompletedTask(observation->taskId);
                operations->DebugRemoveDiagnosticsForTask(observation->taskId);
                state.completedTasks.erase(observation->taskId);
                if (! originalPopup && operations->GetPopupHwndForSelfTest())
                    operations->DebugResetPopupForSelfTest();
            }
        });

        // The reused helper adds only a noninheriting FILE_LIST_DIRECTORY deny.
        // Assume it may have changed security even if its return reports failure.
        cleanable       = false;
        securityChanged = true;
        if (! TryDenyListDirectoryToEveryone(denied))
            return reject(L"actual Windows child DACL denial unavailable");
        WIN32_FIND_DATAW deniedData{};
        wil::unique_hfind deniedListing(
            FindFirstFileExW((denied / L"*").c_str(), FindExInfoBasic, &deniedData, FindExSearchNameMatch, nullptr, FIND_FIRST_EX_LARGE_FETCH));
        const DWORD listingError = GetLastError();
        if (deniedListing || listingError != ERROR_ACCESS_DENIED)
            return reject(L"actual OS listing must return ERROR_ACCESS_DENIED before production admission");
        AppendLog(std::format(L"ACL Copy real pre-admission listing denial: {} error={} originalControl=0x{:04X} originalDaclBytes={}",
                              denied.native(),
                              listingError,
                              static_cast<unsigned int>(originalControl),
                              originalDacl->AclSize));

        // Deliberately omit explicit destinationFileSystem: sourceFileSystem
        // owns native Local source and destination paths, as routine Local Copy.
        const auto admitted = StartFileOperationAndGetId(operations,
                                                         FILESYSTEM_COPY,
                                                         FolderWindow::Pane::Left,
                                                         FolderWindow::Pane::Right,
                                                         sourceFileSystem,
                                                         {source},
                                                         target,
                                                         FILESYSTEM_FLAG_RECURSIVE,
                                                         false,
                                                         0u,
                                                         State::ExecutionMode::PerItem,
                                                         false);
        if (! admitted.has_value())
            return reject(L"normal production recursive Copy admission failed");
        observation->taskId = admitted.value();
        if (auto* task = operations->FindTask(observation->taskId))
            task->_presentationState.store(Task::TaskPresentationState::Presented, std::memory_order_release);
        unsigned int acceptedSkips  = 0u;
        uint64_t acceptedGeneration = 0u;
        bool rejectedPrompt         = false;
        const auto quiet            = RedSalamander::TestSupport::PumpMessagesUntil(
            [&]()
        {
            // Task is borrowed only synchronously in this one predicate; never
            // retain a pointer or prompt reference across a pump/retirement.
            auto* task        = operations->FindTask(observation->taskId);
            const auto prompt = TryGetConflictPromptCopy(task);
            if (task && prompt.has_value() && prompt.value().generation != acceptedGeneration)
            {
                if (acceptedSkips != 0u || prompt.value().bucket != Task::ConflictBucket::AccessDenied ||
                    prompt.value().status != HRESULT_FROM_WIN32(ERROR_ACCESS_DENIED) || prompt.value().sourcePath != denied.native() ||
                    ! PromptHasAction(prompt.value(), Task::ConflictAction::Skip) ||
                    ! task->SubmitConflictDecision(prompt.value().generation, Task::ConflictAction::Skip, false))
                {
                    rejectedPrompt = true;
                    operations->CancelAll();
                }
                else
                {
                    acceptedGeneration = prompt.value().generation;
                    ++acceptedSkips;
                    AppendLog(std::format(L"ACL Copy accepted real AccessDenied Skip generation={} source={}", acceptedGeneration, prompt.value().sourcePath));
                }
            }
            return observation->count == 1u && ! operations->FindTask(observation->taskId) && ! operations->HasActiveOperations();
        },
            {.timeout = SelfTest::Scale(std::chrono::seconds(12)), .operationName = L"ACL Copy real conflict/completion/worker quiet point"});
        AppendLog(std::format(L"ACL Copy observed completion: task={} quiet={} rejectedPrompt={} acceptedSkips={} callbacks={} hr=0x{:08X}",
                              observation->taskId,
                              quiet.conditionMet,
                              rejectedPrompt,
                              acceptedSkips,
                              observation->count,
                              static_cast<uint32_t>(observation->hr)));
        // The provider's child Skip is a partial mutation; the host's typed
        // aggregate reports intentional Skip as S_FALSE, per its result contract.
        if (! quiet.conditionMet || rejectedPrompt || acceptedSkips != 1u || observation->count != 1u || observation->hr != S_FALSE)
        {
            cleanable = false;
            return reject(L"real AccessDenied Skip did not reach exact-once partial Copy completion and worker retirement");
        }
        const auto completed = state.completedTasks.find(observation->taskId);
        if (completed == state.completedTasks.end() || completed->second.hr != S_FALSE || completed->second.configuredMaxConcurrency != 1u ||
            completed->second.sourceItemResults.size() != 1u || completed->second.bridgeSourceDirectoryEnumerationCount != 0u ||
            completed->second.bridgeFileAdmissionCount != 0u || ! completed->second.sourceItemResults[0].has_value())
            return reject(L"real task receipt or serial budget unavailable");
        const auto& result = completed->second.sourceItemResults[0].value();
        AppendLog(std::format(L"ACL Copy observed item: status=0x{:08X} completion={} publication={} source={} stage={} verification={} from={} to={}",
                              static_cast<uint32_t>(result.status),
                              static_cast<unsigned int>(result.completion),
                              static_cast<unsigned int>(result.publication),
                              static_cast<unsigned int>(result.sourceDisposition),
                              static_cast<unsigned int>(result.ownedStageDisposition),
                              static_cast<unsigned int>(result.verification),
                              result.finalSourcePath,
                              result.finalDestinationPath));
        if (result.status != S_FALSE || result.completion != FileOperations::ItemCompletion::Skipped ||
            result.publication != FileOperations::PublicationState::Published || result.sourceDisposition != FileOperations::SourceDisposition::Retained ||
            result.ownedStageDisposition != FileOperations::OwnedStageDisposition::Published ||
            result.verification != FileOperations::VerificationState::NotRequested || result.finalSourcePath != source.native() ||
            result.finalDestinationPath != published.native())
            return reject(L"typed selected-root result lost partial Skip/Published/Retained truth");
        std::vector<State::CompletedTaskSummary> summaries;
        operations->CollectCompletedTasks(summaries);
        const auto summary = std::ranges::find_if(summaries, [&](const auto& value) noexcept { return value.taskId == observation->taskId; });
        if (summary == summaries.end() || summary->resultHr != S_FALSE || summary->publishedItemCount != 1u || summary->retainedSourceCount != 1u ||
            summary->removedSourceCount != 0u || summary->unknownSourceCount != 0u || summary->unknownPublicationItemCount != 0u ||
            summary->indeterminateItemCount != 0u || summary->retainedOwnedStageCount != 0u || summary->unknownOwnedStageCount != 0u)
            return reject(L"completed selected-root axes lost published/retained/partial truth or invented cleanup debt");
        if (! restoreSecurity())
        {
            cleanable = false;
            return reject(L"exact original SD/DACL/protection/control restoration failed; no source traversal or recursive cleanup");
        }
        std::vector<Entry> after;
        auto expected = before;
        expected.push_back(Entry{published.lexically_relative(root).native(), true, {}, {}});
        expected.push_back(Entry{(published / denied.filename()).lexically_relative(root).native(), true, {}, {}});
        const auto accessible = std::ranges::find(before, (source / L"accessible.bin").lexically_relative(root).native(), &Entry::relative);
        if (accessible == before.end())
            return reject(L"accessible reference missing from baseline");
        auto copied     = *accessible;
        copied.relative = (published / L"accessible.bin").lexically_relative(root).native();
        expected.push_back(std::move(copied));
        std::ranges::sort(expected, {}, &Entry::relative);
        if (! capture(true, after) || after != expected)
        {
            cleanable = false;
            return reject(L"closed trees differ: expected every original byte/hash plus published sibling and empty denied destination directory only");
        }
        for (const auto& entry : after)
        {
            AppendLog(std::format(L"ACL Copy closed inventory after: {} kind={} bytes={} BLAKE3={}",
                                  entry.relative,
                                  entry.directory ? L"ordinary-directory" : L"ordinary-file",
                                  entry.bytes.size(),
                                  entry.directory ? std::wstring{} : Utf16FromUtf8ForFileOpsSelfTest(Common::Crypto::EncodeHexDigest(entry.digest))));
        }
        AppendLog(std::format(L"ACL Copy task={}: real AccessDenied Skip; task=S_FALSE; item=S_FALSE/Skipped/Published/Retained; "
                              L"source and reference exact bytes/BLAKE3 unchanged; sibling copied=4099; denied destination child empty; "
                              L"complete before/after entries={}/{}; SD/DACL/control restored exactly; workers retired.",
                              observation->taskId,
                              before.size(),
                              after.size()));
        if (! restoreConfiguration() || ! restoreSettings() || GetForegroundWindow() != foreground)
            return reject(L"provider/prompt/settings/focus restoration failed");
        cleanable = true;
        return true;
    }
    catch (const std::bad_alloc&)
    {
        std::terminate();
    }
    catch (const std::filesystem::filesystem_error&)
    {
        Debug::Error(L"ACL Copy filesystem exception; owned guards drained/restored before failure.");
        return false;
    }
    catch (const std::length_error&)
    {
        Debug::Error(L"ACL Copy bounded storage length failed; owned guards drained/restored before failure.");
        return false;
    }
    catch (const std::format_error&)
    {
        Debug::Error(L"ACL Copy formatting failed; owned guards drained/restored before failure.");
        return false;
    }
}
