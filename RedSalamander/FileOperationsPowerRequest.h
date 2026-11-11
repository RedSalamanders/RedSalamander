#pragma once

#define WIN32_LEAN_AND_MEAN
#define NOMINMAX
#include <windows.h>

#pragma warning(push)
#pragma warning(disable : 4625 4626 5026 5027) // WIL handle owners are intentionally move-only.
#include <wil/resource.h>
#pragma warning(pop)

#include <cstdint>
#include <mutex>

// I63 G7 (D-14): keep the computer awake while a File Operations task runs.
//
// One process-wide Windows power request of type PowerRequestSystemRequired is held while at least one task needs the
// machine awake. The display is never kept on, and closing the lid or choosing Sleep still sleeps: this only asks
// Windows not to start an idle sleep. A failure of the power API is never an error for a task.
namespace FileOperations
{
// What decides whether one task needs the machine awake. All five are read from atomics the task already keeps.
struct KeepAwakeFacts final
{
    bool admitted       = false; // the task holds its operation slot: it is past the queue and its worker is working or about to
    bool finished       = false; // the task reached its terminal state
    bool stopping       = false; // Cancel, Stop or shutdown was requested; rollback and cleanup still run
    bool userPaused     = false; // the user paused the task, or the queue paused it behind a running one
    bool awaitingAnswer = false; // the worker waits for the user's answer and the task has been quiet (AwaitingAnswerIsQuiet)
};

// A worker that waits for an answer does not prove the task is still: other streams of the same task may go on
// transferring. The wait counts as "only waiting" once the task reported no progress for this long. The cost: a task
// that waits for an answer keeps the computer awake for up to this long after its last progress.
inline constexpr uint64_t kAwaitingAnswerQuietMs = 120'000u;

[[nodiscard]] constexpr bool AwaitingAnswerIsQuiet(const uint64_t msSinceLastProgress) noexcept
{
    return msSinceLastProgress >= kAwaitingAnswerQuietMs;
}
static_assert(! AwaitingAnswerIsQuiet(0u));
static_assert(! AwaitingAnswerIsQuiet(kAwaitingAnswerQuietMs - 1u));
static_assert(AwaitingAnswerIsQuiet(kAwaitingAnswerQuietMs));

// A task needs the machine awake when it is admitted, not finished, and either stopping or neither paused nor quietly
// waiting for an answer. A task that is only queued behind another task is not admitted and never counts.
[[nodiscard]] constexpr bool TaskNeedsMachineAwake(const KeepAwakeFacts& facts) noexcept
{
    return facts.admitted && ! facts.finished && (facts.stopping || (! facts.userPaused && ! facts.awaitingAnswer));
}

// The three Win32 calls behind the request, replaceable so self-tests do not depend on the machine's power policy.
// create returns an empty handle when the request cannot be created. set and clear return false on failure.
struct PowerRequestApi final
{
    wil::unique_handle (*create)(void* context, const wchar_t* reason) noexcept = nullptr;
    bool (*set)(void* context, HANDLE request) noexcept                         = nullptr;
    bool (*clear)(void* context, HANDLE request) noexcept                       = nullptr;
    void* context                                                               = nullptr;
};

// Counted holders around one power request. The first Acquire sets PowerRequestSystemRequired, the Release that drops
// the last holder clears it. The request handle is created on the first Acquire and kept until the object is destroyed.
// Thread-safe; its mutex is a leaf (no other lock is taken or any callback made while it is held, apart from the three
// API calls).
class PowerRequest final
{
public:
    struct Counters final
    {
        uint32_t holders            = 0u; // Acquire calls not yet matched by a Release
        uint32_t acquireTransitions = 0u; // changes from no holder to one holder
        uint32_t releaseTransitions = 0u; // changes from one holder to no holder
        uint32_t failures           = 0u; // power API calls that failed
        bool held                   = false; // the system-required request is set right now
    };

    PowerRequest() noexcept;
    explicit PowerRequest(const PowerRequestApi& api) noexcept;
    ~PowerRequest();

    PowerRequest(const PowerRequest&)            = delete;
    PowerRequest(PowerRequest&&)                 = delete;
    PowerRequest& operator=(const PowerRequest&) = delete;
    PowerRequest& operator=(PowerRequest&&)      = delete;

    // The one request of this process, backed by the real Windows API.
    [[nodiscard]] static PowerRequest& Shared() noexcept;

    void Acquire() noexcept;
    // A Release without a matching Acquire is ignored.
    void Release() noexcept;
    [[nodiscard]] Counters GetCounters() const noexcept;

private:
    void SetLocked() noexcept;
    void ClearLocked() noexcept;
    void NoteFailureLocked(const wchar_t* step, DWORD error) noexcept;

    mutable std::mutex _mutex;
    PowerRequestApi _api;
    wil::unique_handle _request;
    uint32_t _holders            = 0u;
    uint32_t _acquireTransitions = 0u;
    uint32_t _releaseTransitions = 0u;
    uint32_t _failures           = 0u;
    bool _held                   = false;
    bool _failureLogged          = false;
};

// One task's share of a PowerRequest: held while TaskNeedsMachineAwake is true for the facts last read, given back when
// it turns false and when the lease is destroyed. Update is idempotent and safe from any thread; it reads the facts
// under the lease's own mutex, so two threads that change different facts of the same task cannot leave a stale answer
// (the thread that locks last reads the newest facts).
class KeepAwakeLease final
{
public:
    explicit KeepAwakeLease(PowerRequest& request) noexcept : _request(&request)
    {
    }
    ~KeepAwakeLease()
    {
        if (_held)
            _request->Release();
    }

    KeepAwakeLease(const KeepAwakeLease&)            = delete;
    KeepAwakeLease(KeepAwakeLease&&)                 = delete;
    KeepAwakeLease& operator=(const KeepAwakeLease&) = delete;
    KeepAwakeLease& operator=(KeepAwakeLease&&)      = delete;

    // gatherFacts is called under the lease's mutex and must return a KeepAwakeFacts without taking other locks.
    template <typename GatherFacts> void Update(GatherFacts&& gatherFacts) noexcept
    {
        std::scoped_lock lock(_mutex);
        const bool wanted = TaskNeedsMachineAwake(gatherFacts());
        if (wanted == _held)
            return;
        _held = wanted;
        if (wanted)
            _request->Acquire();
        else
            _request->Release();
    }

    [[nodiscard]] bool Held() const noexcept
    {
        std::scoped_lock lock(_mutex);
        return _held;
    }

private:
    PowerRequest* _request = nullptr;
    mutable std::mutex _mutex;
    bool _held = false;
};
} // namespace FileOperations
