#include "FolderWindow.FileOperations.Tests.OwnedShutdown.h"
#ifdef ENABLE_TESTS
#pragma warning(push)
#pragma warning(disable : 4625 4626 5026 5027 4514 28182) // Owned mutex/event/COM wrappers, matching selftest ownership policy.
#include "Common.Tests.SandboxPath.h"
#include "ContentDigest.h"
#include "FolderWindow.FileOperationsInternal.h"
#include "Helpers.h"
#include "SelfTest/Common/Common.Tests.Runner.h"
#include "StringConversion.h"
#include "TestSupport/TestSupport.Tests.Runner.h"
#include <TlHelp32.h>
#include <charconv>
#include <cstring>
#include <format>
#include <memory>
#include <mutex>
#include <span>
#include <utility>
#include <wil/com.h>
#include <wil/resource.h>

namespace OwnedShutdownSelfTest
{
namespace
{
struct Pause final
{
    std::mutex mutex;
    wil::unique_event release;
    StageProof proof;
    std::wstring source, destination;
    bool consumed = false;
};
struct Authority final
{
    HWND main       = nullptr;
    DWORD uiTid     = 0u;
    DWORD parentPid = 0u;
    std::filesystem::path root;
    std::wstring nonce;
    std::shared_ptr<Pause> pause;
    uint64_t closeAttempts = 0u, deferred = 0u;
    bool quietAccepted = false, originalsPreserved = false, stopping = false, foregroundPreserved = false;
    uint32_t heartbeatTicks  = 0u;
    uint64_t heartbeatMs     = 0u;
    HRESULT completion       = E_PENDING;
    uint64_t completionCount = 0u;
    HRESULT settingsSave     = E_PENDING;
    bool ownedSettingsSave   = false;
    bool completionAxes      = false;
    std::wstring failure;
};
std::mutex g_mutex;
Authority g_authority;
[[nodiscard]] DWORD PhysicalParentPid() noexcept
{
    const HANDLE raw = CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0u);
    if (raw == INVALID_HANDLE_VALUE)
        return 0u;
    wil::unique_handle snapshot(raw);
    PROCESSENTRY32W entry{};
    entry.dwSize = sizeof(entry);
    if (! snapshot || ! Process32FirstW(snapshot.get(), &entry))
        return 0u;
    do
    {
        if (entry.th32ProcessID == GetCurrentProcessId())
            return entry.th32ParentProcessID;
    } while (Process32NextW(snapshot.get(), &entry));
    return 0u;
}
[[nodiscard]] std::shared_ptr<Pause> PauseCopy() noexcept
{
    std::scoped_lock lock(g_mutex);
    return g_authority.pause;
}
[[nodiscard]] bool WriteReceipt(const std::filesystem::path& path, std::string_view bytes) noexcept
{
    wil::unique_hfile file(
        CreateFileW(path.c_str(), GENERIC_WRITE, FILE_SHARE_READ, nullptr, CREATE_NEW, FILE_ATTRIBUTE_NORMAL | FILE_FLAG_OPEN_REPARSE_POINT, nullptr));
    if (! file || bytes.size() > 64u * 1024u)
        return false;
    DWORD written = 0u;
    return WriteFile(file.get(), bytes.data(), static_cast<DWORD>(bytes.size()), &written, nullptr) && written == bytes.size() && FlushFileBuffers(file.get());
}
} // namespace
HRESULT ArmAuthority(HWND mainWindow,
                     std::wstring_view currentCase,
                     const std::filesystem::path& caseRoot,
                     std::wstring_view runId,
                     std::wstring_view nonce,
                     DWORD parentPid) noexcept
{
    using namespace RedSalamander::TestSupport;
    std::error_code ec;
    const auto base        = ResolveTestSandboxBase(ec);
    const auto expectedRun = base / L"runs" / std::wstring(runId);
    BOOL inJob             = FALSE;
    DWORD ownerPid         = 0u;
    const DWORD ownerTid   = GetWindowThreadProcessId(mainWindow, &ownerPid);
    // A private arm, not an environment-driven switch. The exact owned child must
    // prove its case/mode/run/physical parent/Job and marked no-reparse case path.
    if (currentCase != kCase || SelfTest::GetSelfTestOptions().caseFilter != kCase || GetEnvironmentString(kRole) != L"owned-child-v1" || nonce.empty() ||
        GetEnvironmentString(kNonce) != nonce || GetEnvironmentString(kCaseRoot) != caseRoot.native() || ! IsSafeTestRunId(runId) ||
        GetEnvironmentString(kTestRunIdEnvironmentVariable) != runId || parentPid == 0u || PhysicalParentPid() != parentPid || ec ||
        ! IsProcessInJob(GetCurrentProcess(), nullptr, &inJob) || inJob == FALSE || ownerPid != GetCurrentProcessId() || ownerTid != GetCurrentThreadId() ||
        caseRoot != expectedRun / L"scratch" / L"owned-production-close" || nonce != runId ||
        ! Common::Testing::IsAuthorizedTestSandboxPath(caseRoot, {}, false, ec))
        return E_ACCESSDENIED;
    wil::unique_handle parent(OpenProcess(SYNCHRONIZE, FALSE, parentPid));
    if (! parent || WaitForSingleObject(parent.get(), 0u) != WAIT_TIMEOUT)
        return E_ACCESSDENIED;
    std::scoped_lock lock(g_mutex);
    if (g_authority.main != nullptr)
        return HRESULT_FROM_WIN32(ERROR_BUSY);
    g_authority.main      = mainWindow;
    g_authority.uiTid     = ownerTid;
    g_authority.parentPid = parentPid;
    g_authority.root      = caseRoot;
    g_authority.nonce     = nonce;
    return S_OK;
}
bool UsesProductionRoute(HWND mainWindow) noexcept
{
    std::scoped_lock lock(g_mutex);
    return mainWindow != nullptr && g_authority.main == mainWindow && g_authority.uiTid == GetCurrentThreadId();
}
bool OwnerHasActiveOperations(FolderWindow& owner) noexcept
{
    auto* operations = owner.DebugGetFileOperationState(false);
    return operations && operations->HasActiveOperations();
}
HRESULT ArmStagePause(const std::filesystem::path& source, const std::filesystem::path& destination) noexcept
{
    auto pause = std::make_shared<Pause>();
    pause->release.reset(CreateEventW(nullptr, TRUE, FALSE, nullptr));
    if (! pause->release)
        return HRESULT_FROM_WIN32(GetLastError());
    pause->source      = source.native();
    pause->destination = destination.native();
    std::scoped_lock lock(g_mutex);
    if (g_authority.main == nullptr || g_authority.uiTid != GetCurrentThreadId() || g_authority.pause ||
        ! Common::Testing::IsSameOrDescendantTestSandboxPath(source, g_authority.root) ||
        ! Common::Testing::IsSameOrDescendantTestSandboxPath(destination, g_authority.root))
        return E_ACCESSDENIED;
    g_authority.pause = std::move(pause);
    return S_OK;
}
HRESULT MaybePauseStageAfterWrite(std::wstring_view source,
                                  std::wstring_view destination,
                                  std::wstring_view stage,
                                  IFileSystemBoundObject* ownedStage,
                                  const FileSystemOptions* options,
                                  uint64_t bytes) noexcept
{
    const auto pause = PauseCopy();
    if (! pause)
        return S_OK;
    {
        std::scoped_lock lock(pause->mutex);
        if (pause->consumed || source != pause->source || destination != pause->destination)
            return S_OK;
        pause->consumed = true;
    }
    if (! ownedStage || bytes == 0u)
        return E_UNEXPECTED;
    FileSystemContentProof proof{};
    proof.sizeBytes = sizeof(proof);
    wil::com_ptr<IFileSystemBoundContentProof> contentProof;
    const HRESULT queryHr = ownedStage->QueryInterface(IID_PPV_ARGS(contentProof.addressof()));
    if (FAILED(queryHr) || ! contentProof)
        return FAILED(queryHr) ? queryHr : E_NOINTERFACE;
    const HRESULT proofHr = contentProof->GetContentProof(options, &proof);
    if (FAILED(proofHr) || proof.algorithm != FILESYSTEM_CONTENT_PROOF_BLAKE3_256 || proof.contentSizeBytes != bytes)
        return FAILED(proofHr) ? proofHr : HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
    {
        std::scoped_lock lock(pause->mutex);
        pause->proof.source      = source;
        pause->proof.destination = destination;
        pause->proof.stage       = stage;
        pause->proof.bytes       = bytes;
        std::memcpy(pause->proof.digest.data(), proof.digest, pause->proof.digest.size());
        pause->proof.status  = S_OK;
        pause->proof.entered = true; // publish positive persisted proof last
    }
    // Deliberately hold the worker while cancellation is requested so the child
    // can observe responsive Stopping. Finite bailout; no busy polling or UI wait.
    const DWORD waited    = WaitForSingleObject(pause->release.get(), 15'000u);
    const DWORD waitError = waited == WAIT_FAILED ? GetLastError() : ERROR_SUCCESS;
    {
        std::scoped_lock lock(pause->mutex);
        pause->proof.returned = true;
    }
    return waited == WAIT_OBJECT_0 ? S_OK : waited == WAIT_TIMEOUT ? HRESULT_FROM_WIN32(ERROR_TIMEOUT) : HRESULT_FROM_WIN32(waitError);
}
StageProof GetStageProof()
{
    const auto pause = PauseCopy();
    if (! pause)
        return {};
    std::scoped_lock lock(pause->mutex);
    return pause->proof;
}
void ReleaseStagePause() noexcept
{
    const auto pause = PauseCopy();
    if (pause)
    {
        {
            std::scoped_lock lock(pause->mutex);
            pause->proof.released = true;
        }
        pause->release.SetEvent();
    }
}
void RecordParkedProof(bool originalsPreserved, uint32_t heartbeatTicks, uint64_t heartbeatMs, bool stopping, bool foregroundPreserved) noexcept
{
    std::scoped_lock lock(g_mutex);
    g_authority.originalsPreserved  = originalsPreserved;
    g_authority.heartbeatTicks      = heartbeatTicks;
    g_authority.heartbeatMs         = heartbeatMs;
    g_authority.stopping            = stopping;
    g_authority.foregroundPreserved = foregroundPreserved;
}
void RecordCompletion(HRESULT status, bool exactUnpublishedAndStageRemoved) noexcept
{
    std::scoped_lock lock(g_mutex);
    g_authority.completion = status;
    ++g_authority.completionCount;
    g_authority.completionAxes = exactUnpublishedAndStageRemoved;
}
void RecordFailure(std::wstring_view reason) noexcept
{
    std::scoped_lock lock(g_mutex);
    if (g_authority.failure.empty())
        g_authority.failure = reason;
}
void RecordCloseAttempt(bool active) noexcept
{
    std::scoped_lock lock(g_mutex);
    ++g_authority.closeAttempts;
    if (g_authority.closeAttempts == 1u && ! active)
        g_authority.failure = L"first WM_CLOSE had no parked active task";
}
void RecordDeferredClose(bool active) noexcept
{
    std::scoped_lock lock(g_mutex);
    if (active)
        ++g_authority.deferred;
    else
        g_authority.failure = L"first production close did not defer an active task";
}
void RecordAcceptedQuietClose(bool quiet) noexcept
{
    std::scoped_lock lock(g_mutex);
    g_authority.quietAccepted = quiet;
}
void RecordNormalSettingsSave(bool ownedSave, HRESULT status) noexcept
{
    std::scoped_lock lock(g_mutex);
    g_authority.ownedSettingsSave = ownedSave;
    g_authority.settingsSave      = status;
}
bool RecordNormalDestroy(HWND mainWindow, uint64_t normalTeardownCount) noexcept
{
    const auto stage = GetStageProof();
    std::scoped_lock lock(g_mutex);
    const bool passed = mainWindow == g_authority.main && g_authority.failure.empty() && g_authority.closeAttempts >= 2u && g_authority.deferred >= 1u &&
                        g_authority.quietAccepted && normalTeardownCount == 1u && g_authority.ownedSettingsSave && SUCCEEDED(g_authority.settingsSave) &&
                        g_authority.originalsPreserved && g_authority.stopping && g_authority.heartbeatTicks >= 10u && g_authority.heartbeatMs >= 300u &&
                        g_authority.foregroundPreserved && g_authority.completionCount == 1u && g_authority.completion == HRESULT_FROM_WIN32(ERROR_CANCELLED) &&
                        g_authority.completionAxes && stage.entered && stage.released && stage.returned && stage.status == S_OK && stage.bytes > 0u;
    const auto text = std::format("schema=1\nnonce={}\npid={}\nparentPid={}\npassed={}\nnormalCloseAttempts={}\ndeferredClose={}\nquietBeforeDestroy={}"
                                  "\nnormalTeardown={}\nheartbeatTicks={}\nheartbeatMs={}\ncompletion=0x{:08X}\ncompletionAxes={}\nstageBytes={}\nstageBlake3={"
                                  "}\nfailure={}\nsettingsSave=0x{:08X}\nsettingsSaveOwned={}\ncompletionCount={}\n",
                                  Common::Strings::Utf8FromUtf16ReplacingInvalid(g_authority.nonce),
                                  GetCurrentProcessId(),
                                  g_authority.parentPid,
                                  passed ? 1u : 0u,
                                  g_authority.closeAttempts,
                                  g_authority.deferred,
                                  g_authority.quietAccepted ? 1u : 0u,
                                  normalTeardownCount,
                                  g_authority.heartbeatTicks,
                                  g_authority.heartbeatMs,
                                  static_cast<unsigned long>(g_authority.completion),
                                  g_authority.completionAxes ? 1u : 0u,
                                  stage.bytes,
                                  Common::Crypto::EncodeHexDigest(stage.digest),
                                  Common::Strings::Utf8FromUtf16ReplacingInvalid(g_authority.failure),
                                  static_cast<unsigned long>(g_authority.settingsSave),
                                  g_authority.ownedSettingsSave ? 1u : 0u,
                                  g_authority.completionCount);
    return WriteReceipt(g_authority.root / L"shutdown-receipt.txt", text) && passed;
}
} // namespace OwnedShutdownSelfTest
#pragma warning(pop)
#endif
