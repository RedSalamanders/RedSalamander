// Included from FolderWindow.FileOperations.Tests.Runner.cpp; never compiled standalone.
// F7 baseline only: counters do not alter capability, cache, endpoint or consent policy.
namespace AdmissionBaselineSelfTest
{
constexpr uint32_t kSchema = 1u;
constexpr uint32_t kSlots  = 19u;
using Probe                = HRESULT(__stdcall*)(uint32_t, uint32_t, uint64_t*);
constexpr std::array<std::wstring_view, kSlots> kNames{{L"Schema",
                                                        L"OwnerThreadId",
                                                        L"RouteFactsCalls",
                                                        L"RouteFactsCopyCalls",
                                                        L"RouteFactsMoveCalls",
                                                        L"LegacyPathCapabilitiesCalls",
                                                        L"TransferPeerCalls",
                                                        L"DescriptorCalls",
                                                        L"BuildLocalRouteCalls",
                                                        L"RealVolumePathCalls",
                                                        L"RealVolumeNameCalls",
                                                        L"RootBufferAllocations",
                                                        L"RootBufferRequestedBytes",
                                                        L"RootBufferCapacityBytes",
                                                        L"GuidBufferAllocations",
                                                        L"GuidBufferRequestedBytes",
                                                        L"GuidBufferCapacityBytes",
                                                        L"CacheHits",
                                                        L"CacheMisses"}};

struct InventoryEntry final
{
    size_t root = 0u;
    std::filesystem::path relative;
    bool directory = false;
    std::string bytes;
    Common::Crypto::Blake3Digest digest{};
    bool operator==(const InventoryEntry&) const noexcept = default;
};

// Fixture-specific closed inventory: only marker-owned canonical roots with tiny
// synthetic files, at most 3 levels, no reparse entry of ANY kind. Route aliases are
// not traversed. This differs intentionally from a general consumer file scanner.
[[nodiscard]] bool CaptureInventory(const std::vector<std::filesystem::path>& roots,
                                    const wil::com_ptr<IFileSystemIO>& io,
                                    size_t maxEntries,
                                    std::vector<InventoryEntry>& out) noexcept
{
    out.clear();
    for (size_t rootIndex = 0u; rootIndex < roots.size(); ++rootIndex)
    {
        std::vector<std::pair<std::filesystem::path, unsigned int>> pending{{roots[rootIndex], 0u}};
        for (size_t next = 0u; next < pending.size(); ++next)
        {
            const auto directory   = pending[next].first;
            const auto depth       = pending[next].second;
            const DWORD attributes = GetFileAttributesW(directory.c_str());
            if (attributes == INVALID_FILE_ATTRIBUTES || (attributes & FILE_ATTRIBUTE_REPARSE_POINT) != 0u || (attributes & FILE_ATTRIBUTE_DIRECTORY) == 0u ||
                depth > 3u)
                return false;
            if (next == 0u)
                out.push_back(InventoryEntry{rootIndex, L".", true, {}, {}});
            std::error_code ec;
            std::filesystem::directory_iterator entry(directory, std::filesystem::directory_options::none, ec);
            if (ec)
                return false;
            for (; entry != std::filesystem::directory_iterator{}; entry.increment(ec))
            {
                if (ec || out.size() >= maxEntries)
                    return false;
                const auto path             = entry->path();
                const DWORD entryAttributes = GetFileAttributesW(path.c_str());
                if (entryAttributes == INVALID_FILE_ATTRIBUTES || (entryAttributes & FILE_ATTRIBUTE_REPARSE_POINT) != 0u)
                    return false;
                InventoryEntry value{};
                value.root      = rootIndex;
                value.relative  = path.lexically_relative(roots[rootIndex]);
                value.directory = (entryAttributes & FILE_ATTRIBUTE_DIRECTORY) != 0u;
                if (value.relative.empty() || value.relative.is_absolute())
                    return false;
                if (value.directory)
                    pending.emplace_back(path, depth + 1u);
                else
                {
                    uint64_t size = 0u;
                    if (! GetFileSizeFsIo(io, path, size) || size > 256u || ! ReadFileTextFsIo(io, path, value.bytes) || value.bytes.size() != size)
                        return false;
                    Common::Crypto::Blake3Hasher hasher;
                    hasher.Update(std::as_bytes(std::span<const char>(value.bytes.data(), value.bytes.size())));
                    value.digest = hasher.Finalize();
                }
                out.emplace_back(std::move(value));
            }
            if (ec)
                return false;
        }
    }
    std::ranges::sort(out, [](const InventoryEntry& a, const InventoryEntry& b) noexcept {
        return a.root != b.root ? a.root < b.root : a.relative.native() < b.relative.native();
    });
    return true;
}

struct Gate final
{
    std::atomic<bool> entered{false};
    wil::unique_event release;
};
struct Route final
{
    std::wstring name;
    FileSystemOperation operation = FILESYSTEM_COPY;
    bool sourceUnc                = false;
    bool destinationUnc           = false;
    bool alternateDestination     = false;
    bool mixedSource              = false;
    bool alternateSource          = false;
    bool mountedDestination       = false;
    bool mountedSource            = false; // all source A on the test-owned volume, or that half of a mixed source
};

// Unelevated controls for the handed-off mount gate. The synthetic GUID names no volume and is
// never mounted; the live control observes the primary volume through a plain owned directory.
[[nodiscard]] bool RunOwnedVolumeRefusalControls(const std::filesystem::path& probeDirectory, std::wstring& failure) noexcept
{
    using namespace RedSalamander::TestSupport;
    const std::wstring owned = L"\\\\?\\Volume{00000000-5a4d-4e2b-9c11-7e57c1ea0f00}\\";
    const auto mount         = probeDirectory / L"mount";
    const auto primaryGuid   = TryGetVolumeGuidPathForPath(probeDirectory.root_path());
    if (! SelfTest::EnsureDirectory(probeDirectory) || ! primaryGuid.has_value())
    {
        failure = L"probe directory or primary volume GUID unavailable";
        return false;
    }
    OwnedVirtualDiskMountFacts accepted{};
    accepted.parentAuthorized       = true;
    accepted.directoryIsMountPoint  = true;
    accepted.resolvedVolumeGuidPath = owned;
    accepted.volumePathNames        = {mount.native() + L"\\"};
    accepted.fileBackedVirtualDisk  = true;
    accepted.fileSystemName         = L"NTFS";
    if (const auto refusal = EvaluateOwnedVirtualDiskMount(accepted, owned, mount, mount); ! refusal.empty())
    {
        failure = std::format(L"complete test-owned facts were refused: {}", refusal);
        return false;
    }
    std::wstring operationRefusal;
    const auto ownedChild    = mount / L"case" / L"source";
    const auto operationPath = TryMakeOwnedVirtualDiskOperationPath(accepted, owned, mount, mount, ownedChild, operationRefusal);
    if (! operationPath.has_value() || operationPath.value() != std::filesystem::path(owned) / L"case" / L"source" || ! operationRefusal.empty())
    {
        failure = L"the exact owned child did not resolve to its test-owned volume "
                  L"operation path";
        return false;
    }
    for (const auto& unsafePath : {mount, probeDirectory / L"other", mount / L".." / L"escape", std::filesystem::path(L"relative/child")})
    {
        if (TryMakeOwnedVirtualDiskOperationPath(accepted, owned, mount, mount, unsafePath, operationRefusal).has_value() ||
            operationRefusal != L"operation-path-not-an-owned-child")
        {
            failure = std::format(L"an unowned operation path was accepted: {}", unsafePath.native());
            return false;
        }
    }
    const auto expectRefusal = [&](std::wstring_view control,
                                   const OwnedVirtualDiskMountFacts& facts,
                                   std::wstring_view guid,
                                   const std::filesystem::path& mountPath,
                                   std::wstring_view expectedPrefix) noexcept
    {
        const auto refusal = EvaluateOwnedVirtualDiskMount(facts, guid, mountPath, mount);
        std::wstring operationFailure;
        const bool operationRefused =
            ! TryMakeOwnedVirtualDiskOperationPath(facts, guid, mountPath, mount, mountPath / L"case", operationFailure).has_value() &&
            operationFailure.starts_with(expectedPrefix);
        if (refusal.starts_with(expectedPrefix) && operationRefused)
            return true;
        failure = std::format(L"control {} expected refusal '{}' but got '{}'", control, expectedPrefix, refusal);
        return false;
    };
    auto foreign                   = accepted;
    foreign.resolvedVolumeGuidPath = primaryGuid.value();
    auto lettered                  = accepted;
    lettered.volumePathNames.push_back(L"D:\\");
    auto physical                  = accepted;
    physical.fileBackedVirtualDisk = false;
    auto plain                     = accepted;
    plain.directoryIsMountPoint    = false;
    auto outside                   = accepted;
    outside.parentAuthorized       = false;
    auto fat                       = accepted;
    fat.fileSystemName             = L"FAT32";
    // The real primary volume, as if mounted at the owned path: refused for its
    // drive-letter access path or physical bus, not merely for a GUID mismatch.
    auto realVolume                   = ObserveOwnedVirtualDiskMount(mount, primaryGuid.value());
    const bool liveUnmountedRefused   = expectRefusal(L"live-plain-directory", realVolume, primaryGuid.value(), mount, L"mount-directory-is-not-");
    realVolume.parentAuthorized       = true;
    realVolume.directoryIsMountPoint  = true;
    realVolume.resolvedVolumeGuidPath = primaryGuid.value();
    realVolume.fileSystemName         = L"NTFS";
    return liveUnmountedRefused && expectRefusal(L"foreign-guid", foreign, owned, mount, L"foreign-volume-guid") &&
           expectRefusal(L"drive-letter", lettered, owned, mount, L"volume-reachable-outside-run-mount") &&
           expectRefusal(L"physical-bus", physical, owned, mount, L"volume-not-file-backed-virtual-disk") &&
           expectRefusal(L"not-a-mount-point", plain, owned, mount, L"mount-directory-is-not-") &&
           expectRefusal(L"unmarked-parent", outside, owned, mount, L"mount-parent-not-in-marked") &&
           expectRefusal(L"not-ntfs", fat, owned, mount, L"volume-not-ntfs") &&
           expectRefusal(L"foreign-mount-path", accepted, owned, probeDirectory / L"other", L"mount-path-not-run-owned") &&
           expectRefusal(L"malformed-guid", accepted, L"\\\\?\\Volume{not-a-guid}\\", mount, L"volume-guid-malformed") &&
           (EvaluateOwnedVirtualDiskMount(realVolume, primaryGuid.value(), mount, mount).starts_with(L"volume-not-file-backed-virtual-disk") ||
            expectRefusal(L"live-primary-volume", realVolume, primaryGuid.value(), mount, L"volume-reachable-outside-run-mount"));
}
} // namespace AdmissionBaselineSelfTest

[[nodiscard]] bool RunFileOpsAdmissionBaseline(SelfTestState& state) noexcept
{
    using namespace AdmissionBaselineSelfTest;
    auto* folder      = TryGetFolderWindow(state.mainWindow);
    auto* operations  = TryGetFileOps(folder);
    const auto reject = [&state](std::wstring_view reason) noexcept
    {
        state.retainTempRootOnFailure = true;
        AppendLog(std::format(L"F7 actual admission baseline FAILED: {}", reason));
        return false;
    };
    if (! folder || ! operations || ! CancelOutstandingSelfTestOperations(operations))
        return reject(L"owner/initial quiet point unavailable");
    const auto sourceFs      = folder->GetFileSystem(FolderWindow::Pane::Left);
    const auto destinationFs = folder->GetFileSystem(FolderWindow::Pane::Right);
    wil::com_ptr<IFileSystemIO> io;
    if (! sourceFs || ! destinationFs || ! sourceFs.try_query_to(io.put()) || folder->GetFileSystemPluginId(FolderWindow::Pane::Left) != kPluginIdLocal ||
        folder->GetFileSystemPluginId(FolderWindow::Pane::Right) != kPluginIdLocal)
        return reject(L"actual Local pane providers unavailable");

    RedSalamander::TestSupport::ScopedWindowActivationBlocker blocker;
    if (! RedSalamander::TestSupport::ScopedWindowActivationBlocker::IsActiveForCurrentThread() && ! blocker.Start())
        return reject(L"no-activation guard unavailable");
    RedSalamander::TestSupport::ScopedNonActivatingDesktopObserver desktop;
    if (! desktop.Start())
        return reject(L"process-owned no-activation observation unavailable");
    const auto sourceVtable      = reinterpret_cast<const void* const*>(sourceFs.get())[0];
    const auto destinationVtable = reinterpret_cast<const void* const*>(destinationFs.get())[0];
    auto module                  = AcquireModuleReferenceFromAddress(sourceVtable);
    auto destinationModule       = AcquireModuleReferenceFromAddress(destinationVtable);
    if (! module || ! destinationModule || module.get() != destinationModule.get())
        return reject(L"Local provider module pin mismatch");
    const FARPROC address = GetProcAddress(module.get(), "RedSalamanderFileSystemTestAdmissionProbe");
#pragma warning(push)
#pragma warning(disable : 4191) // Validated POD-only ENABLE_TESTS export; pins outlive all tasks.
    const auto probe = reinterpret_cast<Probe>(address);
#pragma warning(pop)
    if (! probe)
        return reject(L"instrumented Local provider export unavailable");

    const bool autoDismiss         = operations->GetAutoDismissSuccess();
    const auto savedFileOperations = g_settings.fileOperations;
    bool settingsRestored          = false;
    const auto restoreSettings     = wil::scope_exit([&]() noexcept
    {
        if (settingsRestored)
            return;
        operations->SetAutoDismissSuccess(autoDismiss);
        g_settings.fileOperations = savedFileOperations;
        if (! operations->DebugFlushPendingSettingsSaveForSelfTest(5'000u))
            AppendLog(L"F7 cleanup settings save drain failed.");
    });
    operations->SetAutoDismissSuccess(false);
    const auto root = state.tempRoot / L"f7-admission-baseline";
    std::optional<std::filesystem::path> alternateBase;
    std::filesystem::path alternateRoot;
    // Mounted routes use only the test-owned VHDX volume that the elevated harness created,
    // proved and mounted inside this run. This test never mounts, unmounts or adopts any other
    // volume: it validates the handed-off mount and refuses every foreign volume GUID.
    std::filesystem::path mountedRoot;
    std::wstring ownedVolumeGuid;
    bool rootsRemoved            = false;
    const auto cleanupOwnedRoots = [&]() noexcept
    {
        if (rootsRemoved)
            return true;
        if (operations->HasActiveOperations())
        {
            state.retainTempRootOnFailure = true;
            AppendLog(L"F7 cleanup retained owned roots: worker quiet point failed; no filesystem cleanup attempted.");
            return false;
        }
        // Leave a cold ephemeral provider cache after this test, outside every
        // sample. The production one-entry cache/TTL policy remains unchanged.
        std::array<uint64_t, kSlots> ignored{};
        if (FAILED(probe(2u, kSlots, ignored.data())))
            return false;
        if (state.retainTempRootOnFailure)
        {
            // The mounted witness lives on the test-owned VHDX, not under the mount path: after a
            // failed FileOps suite Run-AllTests unmounts and detaches that disk but keeps its file in the run.
            AppendLog(std::format(L"F7 failed admission/storage witness retained: primary={} alternate={} mounted={} "
                                  L"(on the test-owned VHDX, kept detached in this run after a failed FileOps suite)",
                                  root.native(),
                                  alternateRoot.native(),
                                  mountedRoot.native()));
            return false;
        }
        if (! SelfTest::RemoveAll(root))
            return false;
        // Only this case's subtree on the test-owned volume; the harness deletes the disk after a
        // passing FileOps suite and otherwise keeps it detached in the run.
        if (! mountedRoot.empty() && ! SelfTest::RemoveAll(mountedRoot))
            return false;
        if (! alternateRoot.empty())
        {
            if (! SelfTest::RemoveAll(alternateRoot))
                return false;
            std::error_code ec;
            if (std::filesystem::is_empty(alternateBase.value(), ec) && ! ec)
                static_cast<void>(std::filesystem::remove(alternateBase.value(), ec));
            if (ec)
                return false;
            PruneEmptyAlternateVolumeSandboxParents(alternateBase.value());
        }
        rootsRemoved = true;
        return true;
    };
    const auto cleanupRoots = wil::scope_exit([&]() noexcept
    {
        if (! cleanupOwnedRoots())
            AppendLog(L"F7 cleanup of owned roots/cache failed; inspect retained scratch.");
    });
    if (! RecreateEmptyDirectory(root))
        return reject(L"owned primary root unavailable");
    std::wstring refusalFailure;
    if (! RunOwnedVolumeRefusalControls(root / L"owned-volume-refusal-controls", refusalFailure))
        return reject(std::format(L"test-owned volume refusal controls failed: {}", refusalFailure));
    AppendLog(L"F7 OWNED-VOLUME REFUSAL CONTROLS passed: foreign GUID, drive letter, physical bus, plain directory, unmarked parent, "
              L"non-NTFS, foreign mount path, malformed GUID and the live primary volume are refused.");
    std::wstring alternateReason;
    alternateBase = TryCreateAlternateWritableVolumeSelfTestRoot(state.tempRoot, alternateReason);
    alternateRoot = alternateBase.has_value() ? alternateBase.value() / L"f7-admission-baseline" : std::filesystem::path{};
    if (! alternateRoot.empty() && ! RecreateEmptyDirectory(alternateRoot))
        return reject(L"authorized alternate case root unavailable");
    const LoopbackShareStatus uncStatus = ProbeLoopbackShare(root);
    if (uncStatus == LoopbackShareStatus::Refused)
        return reject(L"the sandbox guard refused the case root's loopback administrative-share alias");
    const bool uncAvailable = uncStatus == LoopbackShareStatus::Reachable;
    const auto skip         = [](std::wstring_view route, std::wstring_view reason) noexcept
    {
        AppendLog(std::format(L"F7 route SKIPPED: route={} reason={}; no admission qualification", route, reason));
        Debug::Perf::Emit(L"FileOps.AdmissionBaseline.RouteSkipped", std::format(L"{}:{}", route, reason), 0u, kSchema, 0u, S_FALSE);
    };
    if (! uncAvailable)
        skip(L"UNC/UNC-partitions", L"existing-loopback-admin-share-unreachable");
    if (! alternateBase.has_value())
        skip(L"alternate-volume-managed-Move/volume-partitions", alternateReason);
    // Mounted routes require the harness-provisioned, test-owned VHDX volume. Its absence is a
    // logged skip; a handed-off mount that is not exactly that volume is a failure, never a skip.
    std::wstring mountReason;
    bool mountAvailable = false;
    {
        using namespace RedSalamander::TestSupport;
        const std::wstring handedMount  = GetEnvironmentString(kOwnedVirtualDiskMountEnvironmentVariable);
        const std::wstring handedVolume = GetEnvironmentString(kOwnedVirtualDiskVolumeEnvironmentVariable);
        if (handedMount.empty() && handedVolume.empty())
        {
            const std::wstring harnessReason = GetEnvironmentString(kOwnedVirtualDiskUnavailableEnvironmentVariable);
            mountReason =
                std::format(L"test-owned-vhdx-not-provisioned:{}", harnessReason.empty() ? L"requires-an-elevated-Run-AllTests" : harnessReason.c_str());
        }
        else
        {
            const std::wstring testRoot = GetEnvironmentString(L"REDSALAMANDER_TEST_ROOT");
            const std::wstring runId    = GetEnvironmentString(L"REDSALAMANDER_TEST_RUN_ID");
            const auto expectedMount = testRoot.empty() || ! IsSafeTestRunId(runId) ? std::filesystem::path{}
                                                                                    : GetOwnedVirtualDiskMountPath(testRoot, runId, L"f7-admission-baseline");
            const std::filesystem::path mountPath(handedMount);
            const std::wstring refusal =
                EvaluateOwnedVirtualDiskMount(ObserveOwnedVirtualDiskMount(mountPath, handedVolume), handedVolume, mountPath, expectedMount);
            if (! refusal.empty())
                return reject(std::format(L"handed-off mount refused, no route uses it: {}", refusal));
            mountedRoot = mountPath / L"f7-admission-baseline";
            if (! RecreateEmptyDirectory(mountedRoot))
                return reject(L"test-owned mounted case root unavailable");
            ownedVolumeGuid = handedVolume;
            mountAvailable  = true;
            AppendLog(std::format(L"F7 TEST-OWNED VOLUME: mount={} volumeGuid={} caseRoot={}", mountPath.native(), ownedVolumeGuid, mountedRoot.native()));
        }
    }
    if (! mountAvailable)
        skip(L"test-owned-mounted-volume", mountReason);

    std::vector<Route> routes{{L"local-copy", FILESYSTEM_COPY}};
    if (alternateBase.has_value())
    {
        routes.push_back({L"alternate-managed-move", FILESYSTEM_MOVE, false, false, true});
        routes.push_back({L"volume-partitions-copy", FILESYSTEM_COPY, false, false, false, true, true});
        routes.push_back({L"volume-partitions-move-mixed", FILESYSTEM_MOVE, false, false, false, true, true});
    }
    if (mountAvailable)
    {
        routes.push_back({L"local-to-mounted-copy", FILESYSTEM_COPY, false, false, true, false, false, true});
        routes.push_back({L"local-to-mounted-managed-move", FILESYSTEM_MOVE, false, false, true, false, false, true});
        routes.push_back({L"mounted-to-local-copy", FILESYSTEM_COPY, false, false, false, false, false, false, true});
        routes.push_back({L"mounted-to-local-managed-move", FILESYSTEM_MOVE, false, false, false, false, false, false, true});
        routes.push_back({L"mounted-volume-partitions-copy", FILESYSTEM_COPY, false, false, false, true, true, false, true});
        routes.push_back({L"mounted-volume-partitions-move-mixed", FILESYSTEM_MOVE, false, false, false, true, true, false, true});
    }
    if (uncAvailable)
    {
        routes.push_back({L"local-to-unc-copy", FILESYSTEM_COPY, false, true});
        routes.push_back({L"local-to-unc-managed-move", FILESYSTEM_MOVE, false, true});
        routes.push_back({L"unc-to-local-copy", FILESYSTEM_COPY, true});
        routes.push_back({L"unc-to-local-managed-move", FILESYSTEM_MOVE, true});
        routes.push_back({L"unc-partitions-copy", FILESYSTEM_COPY, false, false, false, true});
        routes.push_back({L"unc-partitions-move-mixed", FILESYSTEM_MOVE, false, false, false, true});
    }

    uint64_t sampleSerial = 0u;
    for (const auto& route : routes)
        for (const size_t selectedCount : std::array<size_t, 3>{{1u, 128u, 4096u}})
            for (const bool alternating : std::array<bool, 2>{{false, true}})
            {
                if (alternating && ! route.mixedSource)
                    continue;
                const bool mountedRoute           = route.mountedDestination || route.mountedSource;
                const bool alternatePrimarySource = route.mountedSource && ! route.mixedSource;
                const bool hasAlternate           = route.alternateDestination || route.alternateSource || alternatePrimarySource;
                const auto localCase              = root / L"case";
                // The second root is the D fixture, or the test-owned volume for mounted routes.
                const auto secondRoot    = mountedRoute ? mountedRoot : alternateRoot;
                const auto alternateCase = secondRoot.empty() ? std::filesystem::path{} : secondRoot / L"case";
                if (! RecreateEmptyDirectory(localCase) || (hasAlternate && ! RecreateEmptyDirectory(alternateCase)))
                    return reject(L"case directories unavailable");
                const auto sourceA     = (alternatePrimarySource ? alternateCase : localCase) / L"source-a";
                const auto sourceB     = route.alternateSource ? alternateCase / L"source-b" : localCase / L"source-b";
                const auto destination = route.alternateDestination ? alternateCase / L"destination" : localCase / L"destination";
                if (! SelfTest::EnsureDirectory(sourceA) || ! SelfTest::EnsureDirectory(sourceB) || ! SelfTest::EnsureDirectory(destination) ||
                    ! WriteFileTextFsIo(io, destination / L"unrelated-sentinel.bin", "F7 destination sentinel; preserve exactly"))
                    return reject(L"source/destination staging unavailable");
                std::vector<std::filesystem::path> inventoryRoots{localCase};
                if (hasAlternate)
                    inventoryRoots.push_back(alternateCase);
                std::vector<InventoryEntry> expectedInventory;
                const auto addSeed = [&](size_t rootIndex, const std::filesystem::path& path, bool directory, std::string_view bytes) noexcept
                {
                    InventoryEntry entry{};
                    entry.root      = rootIndex;
                    entry.relative  = path == inventoryRoots[rootIndex] ? std::filesystem::path(L".") : path.lexically_relative(inventoryRoots[rootIndex]);
                    entry.directory = directory;
                    entry.bytes     = bytes;
                    if (! directory)
                    {
                        Common::Crypto::Blake3Hasher hasher;
                        hasher.Update(std::as_bytes(std::span<const char>(bytes.data(), bytes.size())));
                        entry.digest = hasher.Finalize();
                    }
                    expectedInventory.emplace_back(std::move(entry));
                };
                for (size_t rootIndex = 0u; rootIndex < inventoryRoots.size(); ++rootIndex)
                    addSeed(rootIndex, inventoryRoots[rootIndex], true, {});
                addSeed(alternatePrimarySource ? 1u : 0u, sourceA, true, {});
                addSeed(route.alternateSource ? 1u : 0u, sourceB, true, {});
                addSeed(route.alternateDestination ? 1u : 0u, destination, true, {});
                addSeed(route.alternateDestination ? 1u : 0u, destination / L"unrelated-sentinel.bin", false, "F7 destination sentinel; preserve exactly");
                std::vector<std::filesystem::path> selectedPaths;
                selectedPaths.reserve(selectedCount);
                for (size_t index = 0u; index < selectedCount; ++index)
                {
                    const bool second = route.mixedSource && (alternating ? index % 2u != 0u : index >= (selectedCount + 1u) / 2u);
                    const auto source = (second ? sourceB : sourceA) / std::format(L"item-{:04}.bin", index);
                    const std::string bytes =
                        std::format("F7|op={}|n={}|i={}|distinct original payload", static_cast<unsigned int>(route.operation), selectedCount, index);
                    if (! WriteFileTextFsIo(io, source, bytes))
                        return reject(L"distinct source payload staging failed");
                    addSeed(alternatePrimarySource || (second && route.alternateSource) ? 1u : 0u, source, false, bytes);
                    const bool alias                      = route.sourceUnc || (route.mixedSource && ! route.alternateSource && second);
                    std::filesystem::path admissionSource = alias ? LoopbackShareAlias(source) : source;
                    selectedPaths.push_back(std::move(admissionSource));
                }
                auto providerDestination = route.destinationUnc ? LoopbackShareAlias(destination) : destination;
                std::ranges::sort(expectedInventory, [](const InventoryEntry& a, const InventoryEntry& b) noexcept {
                    return a.root != b.root ? a.root < b.root : a.relative.native() < b.relative.native();
                });
                std::vector<InventoryEntry> beforeInventory;
                if (! CaptureInventory(inventoryRoots, io, selectedCount + 16u, beforeInventory) || beforeInventory != expectedInventory)
                    return reject(L"closed original inventory/bytes/Blake3 differs from the exact seed");
                Common::Crypto::Blake3Hasher inventoryHasher;
                for (const auto& entry : beforeInventory)
                {
                    // Versioned Windows-only receipt: root index, UTF-16 name with
                    // NUL delimiter, kind, byte length and canonical per-file digest.
                    const uint64_t rootIndex  = entry.root;
                    const uint64_t byteLength = entry.bytes.size();
                    const wchar_t delimiter   = L'\0';
                    const uint8_t kind        = entry.directory ? 1u : 0u;
                    inventoryHasher.Update(std::as_bytes(std::span(&rootIndex, 1u)));
                    inventoryHasher.Update(std::as_bytes(std::span(entry.relative.native().data(), entry.relative.native().size())));
                    inventoryHasher.Update(std::as_bytes(std::span(&delimiter, 1u)));
                    inventoryHasher.Update(std::as_bytes(std::span(&kind, 1u)));
                    inventoryHasher.Update(std::as_bytes(std::span(&byteLength, 1u)));
                    inventoryHasher.Update(entry.digest);
                }
                const auto inventoryDigest = Utf16FromUtf8ForFileOpsSelfTest(Common::Crypto::EncodeHexDigest(inventoryHasher.Finalize()));
                FileSystemRouteContract::Snapshot directCFacts;
                FileSystemRouteContract::Snapshot directDFacts; // the test-owned volume's endpoint on mounted routes
                if (mountedRoute)
                {
                    // Seed and inventories use the owned C case and the case on the test-owned
                    // volume, reached only through its run-owned mount; recheck the actual volume.
                    const auto cFacts      = FileSystemRouteContract::Query(sourceFs.get(), localCase.native(), route.operation, kPluginIdLocal);
                    const auto mFacts      = FileSystemRouteContract::Query(sourceFs.get(), alternateCase.native(), route.operation, kPluginIdLocal);
                    const auto mountedGuid = TryGetVolumeGuidPathForPath(alternateCase);
                    if (! mountedGuid.has_value() || _wcsicmp(mountedGuid.value().c_str(), ownedVolumeGuid.c_str()) != 0)
                        return reject(L"mounted case no longer resolves to the test-owned volume GUID");
                    if (cFacts.state != FileSystemRouteContract::QueryState::Available || mFacts.state != FileSystemRouteContract::QueryState::Available ||
                        cFacts.snapshot.rootId.empty() || mFacts.snapshot.rootId.empty() || cFacts.snapshot.rootId == mFacts.snapshot.rootId ||
                        cFacts.snapshot.pathProfileId != mFacts.snapshot.pathProfileId)
                        return reject(L"actual mounted path must resolve to the test-owned volume endpoint, distinct from C");
                    directCFacts = cFacts.snapshot;
                    directDFacts = mFacts.snapshot;
                    AppendLog(std::format(L"F7 TEST-OWNED MOUNT ROUTE: route={} n={} volumeGuid={} mountedCase={} C-root={} mounted-root={} "
                                          L"sameLocalProfile={} inventories=owned-C-and-owned-volume-case-only",
                                          route.name,
                                          selectedCount,
                                          mountedGuid.value(),
                                          alternateCase.native(),
                                          directCFacts.rootId,
                                          directDFacts.rootId,
                                          directDFacts.pathProfileId));
                }

                // A complete same-route admission/cancel/drain primes the warm sample;
                // elapsed prime age is retained, never assumed to fit the 2-second TTL.
                for (unsigned int repetition = 0u; repetition < 3u; ++repetition)
                {
                    for (unsigned int phase = 0u; phase < 3u; ++phase)
                    {
                        const bool prime = phase == 1u;
                        const bool warm  = phase == 2u;
                        const auto tag   = std::format(L"v{}|sample={}|route={}|op={}|n={}|order={}|rep={}|cache={}",
                                                       kSchema,
                                                       ++sampleSerial,
                                                       route.name,
                                                       static_cast<unsigned int>(route.operation),
                                                       selectedCount,
                                                       alternating ? L"alternating" : L"clustered",
                                                       repetition,
                                                       prime  ? L"prime"
                                                       : warm ? L"warm-requested"
                                                              : L"cold-at-ingress");
                        if (operations->HasActiveOperations())
                            return reject(L"cache preparation requires a quiet owner");
                        std::array<uint64_t, kSlots> counts{};
                        if (! warm && FAILED(probe(2u, kSlots, counts.data())))
                            return reject(L"cold cache reset failed");
                        // Action 4 reads the existing cache's last resolution timestamp
                        // into slot 2 outside TLS/timing; it does not change cache policy.
                        if (warm && FAILED(probe(4u, kSlots, counts.data())))
                            return reject(L"warm cache age snapshot failed");
                        const uint64_t primeAgeMs = warm ? GetTickCount64() - counts[2] : 0u;
                        auto gate                 = std::make_shared<Gate>();
                        gate->release.create();
                        std::function<HRESULT()> decisionGate = [gate]() noexcept
                        {
                            gate->entered.store(true, std::memory_order_release);
                            const DWORD wait = WaitForSingleObject(gate->release.get(), 120'000u);
                            return wait == WAIT_OBJECT_0 ? S_OK : wait == WAIT_TIMEOUT ? HRESULT_FROM_WIN32(ERROR_TIMEOUT) : HRESULT_FROM_WIN32(GetLastError());
                        };
                        uint64_t taskId        = 0u;
                        unsigned int completed = 0u;
                        HRESULT result         = E_PENDING;
                        const uint64_t subscription =
                            folder->AddFileOperationCompletedCallback([&](const FolderWindow::FileOperationCompletedEvent& event) noexcept
                        {
                            if (event.taskId == taskId)
                            {
                                ++completed; // completion callback is owner-thread only
                                result = event.hr;
                            }
                        });
                        const auto unsubscribe = wil::scope_exit([&]() noexcept { folder->RemoveFileOperationCompletedCallback(subscription); });
                        // This scope executes before unsubscribe and before borrowed callback
                        // variables/module pins/fixture roots expire, including all failures.
                        const auto drain          = wil::scope_exit([&]() noexcept
                        {
                            if (auto* pending = operations->FindTask(taskId))
                                pending->RequestCancel();
                            gate->release.SetEvent();
                            const bool cancelled = CancelOutstandingSelfTestOperations(operations);
                            const auto retired   = RedSalamander::TestSupport::PumpMessagesUntil([&]() noexcept {
                                return ! operations->HasActiveOperations() && (taskId == 0u || ! operations->FindTask(taskId));
                            }, {.timeout = SelfTest::Scale(std::chrono::seconds(8)), .operationName = L"F7 failure cleanup task retirement"});
                            if (! cancelled || ! retired.conditionMet)
                            {
                                AppendLog(L"F7 failure cleanup hard-joins the isolated owner before borrowed callback/environment state unwinds.");
                                operations->Shutdown();
                            }
                            if (taskId != 0u)
                            {
                                operations->DismissCompletedTask(taskId);
                                operations->DebugRemoveDiagnosticsForTask(taskId);
                            }
                            operations->DebugResetPopupForSelfTest();
                        });
                        auto admissionPaths       = selectedPaths; // argument preparation excluded
                        auto admissionDestination = providerDestination;
                        bool tlsEnabled           = false;
                        const auto endTls         = wil::scope_exit([&]() noexcept
                        {
                            if (tlsEnabled)
                                static_cast<void>(probe(1u, kSlots, counts.data()));
                        });
                        if (FAILED(probe(0u, kSlots, counts.data())))
                            return reject(L"TLS begin failed");
                        tlsEnabled                = true;
                        const uint64_t startedUs  = FolderWindowFileOperationsStateInternal::PerfNowUs();
                        const HRESULT admissionHr = operations->AdmitOperation(route.operation,
                                                                               FolderWindow::Pane::Left,
                                                                               FolderWindow::Pane::Right,
                                                                               sourceFs,
                                                                               std::move(admissionPaths),
                                                                               std::move(admissionDestination),
                                                                               FILESYSTEM_FLAG_NONE,
                                                                               false,
                                                                               0u,
                                                                               FolderWindow::FileOperationState::ExecutionMode::PerItem,
                                                                               false,
                                                                               destinationFs,
                                                                               &taskId,
                                                                               {},
                                                                               {},
                                                                               {},
                                                                               {},
                                                                               std::nullopt,
                                                                               FileOperations::DeleteOrigin::PaneCommand,
                                                                               std::nullopt,
                                                                               std::nullopt,
                                                                               {},
                                                                               std::move(decisionGate));
                        const uint64_t wallUs     = FolderWindowFileOperationsStateInternal::PerfElapsedUs(startedUs);
                        const HRESULT snapshotHr  = probe(1u, kSlots, counts.data());
                        tlsEnabled                = false;
                        if (FAILED(snapshotHr) || counts[0] != kSchema || counts[1] != GetCurrentThreadId() || admissionHr != S_OK || taskId == 0u)
                            return reject(std::format(L"{} admission=0x{:08X} snapshot=0x{:08X} id={}",
                                                      tag,
                                                      static_cast<unsigned long>(admissionHr),
                                                      static_cast<unsigned long>(snapshotHr),
                                                      taskId));
                        // An accepted ingress is not prepared-task or preservation proof.
                        // Retain the actual UI cost/counters even if worker preparation fails.
                        Debug::Perf::Emit(L"FileOps.AdmissionBaseline.IngressUiWallUs", tag, wallUs, selectedCount, taskId, admissionHr);
                        AppendLog(std::format(L"F7 INGRESS {} task={} wallUs={} rootQueries={} rootScratchAllocations={} rootScratchBytes={} "
                                              L"guidQueries={} cacheHits={} cacheMisses={} qualification=pending-preparation-and-storage",
                                              tag,
                                              taskId,
                                              wallUs,
                                              counts[9],
                                              counts[11],
                                              counts[12],
                                              counts[10],
                                              counts[17],
                                              counts[18]));
                        const auto entered = RedSalamander::TestSupport::PumpMessagesUntil([&]() noexcept {
                            return gate->entered.load(std::memory_order_acquire);
                        }, {.timeout = SelfTest::Scale(std::chrono::seconds(60)), .operationName = L"F7 actual admission preconsumption gate"});
                        if (! entered.conditionMet)
                            return reject(entered.timeoutDiagnostic);
                        uint32_t planCount    = 0u;
                        uint32_t strategyMask = 0u;
                        if (! DebugGetPreparedTransferPlanShapeForSelfTest(taskId, &planCount, &strategyMask))
                            return reject(L"actual prepared plan shape missing");
                        const uint32_t expectedPlans = route.mixedSource && selectedCount > 1u ? 2u : 1u;
                        const uint32_t managedBit    = 1u << static_cast<unsigned int>(FileOperations::OperationStrategy::Managed);
                        const uint32_t copyBit       = 1u << static_cast<unsigned int>(FileOperations::OperationStrategy::Copy);
                        const uint32_t nativeBit     = 1u << static_cast<unsigned int>(FileOperations::OperationStrategy::Native);
                        const uint32_t expectedMask  = route.operation == FILESYSTEM_COPY ? copyBit
                                                       : route.mixedSource                ? (selectedCount > 1u ? managedBit | nativeBit : nativeBit)
                                                                                          : managedBit;
                        if (planCount != expectedPlans || strategyMask != expectedMask)
                            return reject(std::format(
                                L"{} actual plans={} mask={} expected plans={} mask={}", tag, planCount, strategyMask, expectedPlans, expectedMask));
                        auto* task = operations->FindTask(taskId);
                        if (! task)
                            return reject(L"held actual task missing");
                        const auto plans = task->LoadPlans();
                        if (! plans)
                            return reject(L"held immutable plans missing");
                        size_t cSelected = 0u;
                        size_t dSelected = 0u;
                        for (const auto& plan : *plans)
                        {
                            const auto* transfer = std::get_if<FileOperations::TransferPlan>(&plan);
                            if (! transfer)
                                return reject(L"admitted transfer has unexpected plan kind");
                            if (mountedRoute)
                            {
                                const auto& expectedDestination = route.mountedDestination ? directDFacts : directCFacts;
                                if (transfer->destinationEndpoint.rootId != expectedDestination.rootId ||
                                    transfer->destinationEndpoint.profileId != expectedDestination.pathProfileId ||
                                    transfer->sourceEndpoint.profileId != directCFacts.pathProfileId)
                                    return reject(L"actual mounted transfer plan endpoint qualification differs");
                                if (transfer->sourceEndpoint.rootId == directCFacts.rootId)
                                    cSelected += transfer->selectedItems.size();
                                else if (transfer->sourceEndpoint.rootId == directDFacts.rootId)
                                    dSelected += transfer->selectedItems.size();
                                else
                                    return reject(L"mounted plan contains an unexpected source root");
                                const auto expectedStrategy = route.operation == FILESYSTEM_COPY ? FileOperations::OperationStrategy::Copy
                                                              : transfer->sourceEndpoint.rootId == transfer->destinationEndpoint.rootId
                                                                  ? FileOperations::OperationStrategy::Native
                                                                  : FileOperations::OperationStrategy::Managed;
                                if (transfer->strategy != expectedStrategy)
                                    return reject(L"mounted partition strategy differs from its actual source/destination roots");
                            }
                            AppendLog(
                                std::format(L"F7 endpoint {} sourceProfile={} sourceRoot={} destinationProfile={} destinationRoot={} strategy={} selected={}",
                                            tag,
                                            transfer->sourceEndpoint.profileId,
                                            transfer->sourceEndpoint.rootId,
                                            transfer->destinationEndpoint.profileId,
                                            transfer->destinationEndpoint.rootId,
                                            static_cast<unsigned int>(transfer->strategy),
                                            transfer->selectedItems.size()));
                        }
                        if (mountedRoute)
                        {
                            const size_t expectedD = alternatePrimarySource ? selectedCount : route.mixedSource ? selectedCount / 2u : 0u;
                            if (dSelected != expectedD || cSelected != selectedCount - expectedD)
                                return reject(L"mounted source partitions lost/reassigned selected items");
                            AppendLog(std::format(L"F7 MOUNT PARTITIONS: {} C-selected={} D-selected={} expectedD={} actualPlans={}",
                                                  tag,
                                                  cSelected,
                                                  dSelected,
                                                  expectedD,
                                                  planCount));
                        }
                        task->RequestCancel();
                        task = nullptr; // no Task borrow crosses the completion pump
                        gate->release.SetEvent();
                        const auto quiet = RedSalamander::TestSupport::PumpMessagesUntil([&]() noexcept {
                            return completed == 1u && ! operations->FindTask(taskId) && ! operations->HasActiveOperations();
                        }, {.timeout = SelfTest::Scale(std::chrono::seconds(60)), .operationName = L"F7 canceled admission worker quiet point"});
                        if (! quiet.conditionMet || result != HRESULT_FROM_WIN32(ERROR_CANCELLED))
                            return reject(std::format(L"{} canceled quiet={} completed={} result=0x{:08X} diagnostic={}",
                                                      tag,
                                                      quiet.conditionMet,
                                                      completed,
                                                      static_cast<unsigned long>(result),
                                                      quiet.timeoutDiagnostic));
                        std::array<uint64_t, kSlots> afterWorker{};
                        if (FAILED(probe(3u, kSlots, afterWorker.data())) || afterWorker != counts)
                            return reject(L"independent provider worker work polluted UI TLS sample");
                        // Prime+warm readback is combined AFTER both known preconsumption
                        // cancellations. Reading 4096 files between them would age the TTL.
                        if (! prime)
                        {
                            std::vector<InventoryEntry> afterInventory;
                            // Workers are already quiet: inspect once and fail on any
                            // mismatch rather than retrying until changed bytes disappear.
                            if (! CaptureInventory(inventoryRoots, io, selectedCount + 16u, afterInventory))
                                return reject(L"post-quiet closed no-reparse inventory could not be observed");
                            if (afterInventory != beforeInventory)
                                return reject(L"canceled admission changed closed inventory/exact bytes/Blake3");
                        }
                        // Desktop observations stay outside the measured admission span and closed storage inventory.
                        if (! desktop.Sample())
                            return reject(std::format(L"{} owned foreground or unavailable desktop observation (failure={}, ownedTransitions={}); "
                                                      L"storage inventory is a separate predicate",
                                                      tag,
                                                      static_cast<unsigned>(desktop.Failure()),
                                                      desktop.OwnedForegroundTransitions()));
                        const auto actualSource      = folder->GetFileSystem(FolderWindow::Pane::Left);
                        const auto actualDestination = folder->GetFileSystem(FolderWindow::Pane::Right);
                        if (actualSource.get() != sourceFs.get() || actualDestination.get() != destinationFs.get())
                            return reject(
                                std::format(L"{} actual pane provider identity changed source={} expectedSource={} destination={} expectedDestination={}",
                                            tag,
                                            reinterpret_cast<uintptr_t>(actualSource.get()),
                                            reinterpret_cast<uintptr_t>(sourceFs.get()),
                                            reinterpret_cast<uintptr_t>(actualDestination.get()),
                                            reinterpret_cast<uintptr_t>(destinationFs.get())));
                        if (! prime)
                        {
                            Debug::Perf::Emit(L"FileOps.AdmissionBaseline.UiWallUs", tag, wallUs, selectedCount, planCount, admissionHr);
                            Debug::Perf::Emit(L"FileOps.AdmissionBaseline.PrimeAgeMs", tag, 0u, primeAgeMs, strategyMask, admissionHr);
                            for (size_t slot = 0u; slot < counts.size(); ++slot)
                                Debug::Perf::Emit(
                                    std::format(L"FileOps.AdmissionBaseline.{}", kNames[slot]), tag, 0u, counts[slot], selectedCount, admissionHr);
                            AppendLog(std::format(L"F7 SAMPLE {} task={} wallUs={} plans={} mask={} primeAgeMs={} cacheState={} inventoryEntries={} "
                                                  L"inventoryBlake3={} exactBytesAndBlake3=preserved foreground=preserved",
                                                  tag,
                                                  taskId,
                                                  wallUs,
                                                  planCount,
                                                  strategyMask,
                                                  primeAgeMs,
                                                  warm && primeAgeMs > 2'000u ? L"prime-expired-at-ingress"
                                                  : warm                      ? L"recent-single-cache-entry-at-ingress; see per-call hits/misses"
                                                                              : L"cold-at-ingress",
                                                  beforeInventory.size(),
                                                  inventoryDigest));
                        }
                    }
                }
            }
    if (! cleanupOwnedRoots())
        return reject(L"final owned-root/cache cleanup failed");
    operations->SetAutoDismissSuccess(autoDismiss);
    g_settings.fileOperations = savedFileOperations;
    if (! operations->DebugFlushPendingSettingsSaveForSelfTest(5'000u))
        return reject(L"final settings save drain failed");
    settingsRestored = true;
    if (! desktop.Finish())
        return reject(std::format(L"final cleanup owned foreground or unavailable event drain (failure={}, ownedTransitions={})",
                                  static_cast<unsigned>(desktop.Failure()),
                                  desktop.OwnedForegroundTransitions()));
    const auto& desktopStats = desktop.Stats();
    AppendLog(std::format(L"F7 no-activation evidence: samples={} ownedTransitions={} foreignForegroundChanges={} cursorChanges={} inputChanges={} drainUs={}",
                          desktopStats.samples,
                          desktop.OwnedForegroundTransitions(),
                          desktopStats.foregroundChanges,
                          desktopStats.cursorChanges,
                          desktopStats.inputChanges,
                          desktopStats.barrierUs));
    return true;
}
