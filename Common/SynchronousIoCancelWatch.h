#pragma once

// Synchronous Win32 I/O cancel watch (plan R0f-SMB).
//
// A thread that makes a synchronous Win32 call on a network path (SMB open, enumerate, write,
// rename, delete, flush, close) has no cooperative checkpoint while the call is pending: a dead
// share holds it until the redirector gives up. Windows can cancel that pending call from another
// thread with CancelSynchronousIo, after which the call fails with ERROR_OPERATION_ABORTED.
//
// A Scope registers the calling thread together with a "should cancel" predicate. One threadpool
// timer per module polls every registered scope every kPollIntervalMs. Once a predicate has been
// true for kGraceMs (time for ordinary cooperative checkpoints and owned-stage cleanup to finish on
// their own), the watch calls CancelSynchronousIo for that thread on every poll until the scope
// ends. Nothing is abandoned or terminated: the blocked call returns to its caller with the abort
// status, and the caller reports it like any other failure under a requested cancel.
//
// The watch is per module (the host and each plugin that includes this header own their own timer
// and registry) so a plugin never depends on host internals and unloads cleanly: the registry is
// empty whenever no scope is alive, and the timer is stopped and drained before the module's
// statics go away.
//
// Owned-stage cleanup budget (I35 R7.4). Removing an exact object the operation already owns (an
// AbortOwnedObject, the raw cleanup of a just-created stage, a publication rollback that restores the
// user's original, the backup removal after a publication) runs under its own short control, whose
// deadline is FILESYSTEM_OWNED_STAGE_CLEANUP_TIMEOUT_MS ahead. A CleanupWindow opened around that I/O
// tells every module's watch to leave the thread's synchronous calls alone until that deadline; after
// it, a dead share is canceled as before. The window is one deadline per thread shared by the whole
// process (it lives in Common.dll), because a provider's cleanup often runs on a host worker thread
// that only the host's watch registered, and a timer thread cannot read another thread's TLS: each
// registration keeps the address of its thread's deadline and the watch reads it there.

#include <atomic>
#include <mutex>
#include <utility>
#include <vector>

#define WIN32_LEAN_AND_MEAN
#define NOMINMAX
#include <Windows.h>

#pragma warning(push)
#pragma warning(disable : 4625 4626 5026 5027 4820 28182) // WIL wrappers are intentionally move-only.
#include <wil/resource.h>
#pragma warning(pop)

// DLL export/import macro for Common.dll. COMMON_EXPORTS is defined by the Common project.
#ifndef COMMON_API
#ifdef COMMON_EXPORTS
#define COMMON_API __declspec(dllexport)
#else
#define COMMON_API __declspec(dllimport)
#endif
#endif

namespace Common::SynchronousIoCancelWatch
{
inline constexpr DWORD kPollIntervalMs = 50u;
inline constexpr DWORD kGraceMs        = 500u;
// The longest cleanup window: FILESYSTEM_OWNED_STAGE_CLEANUP_TIMEOUT_MS (static_assert at its users).
inline constexpr ULONGLONG kMaxCleanupWindowMs = 2'000u;

// Returns true when the registered thread's pending synchronous call must be canceled.
using ShouldCancelFn = bool (*)(void* context) noexcept;

// The calling thread's owned-stage cleanup deadline (GetTickCount64; 0 = none), one per thread for the
// whole process. Valid while the thread runs; a registration reads it only while its scope is alive.
[[nodiscard]] COMMON_API std::atomic<ULONGLONG>* CurrentThreadCleanupDeadline() noexcept;

namespace Detail
{
using UniqueThreadpoolTimer = wil::unique_any<PTP_TIMER, decltype(&::CloseThreadpoolTimer), ::CloseThreadpoolTimer>;

struct Node final
{
    Node() noexcept              = default;
    Node(const Node&)            = delete;
    Node& operator=(const Node&) = delete;
    Node(Node&&)                 = delete;
    Node& operator=(Node&&)      = delete;

    wil::unique_handle thread;
    ShouldCancelFn shouldCancel = nullptr;
    void* context               = nullptr;
    ULONGLONG cancelSinceTick   = 0u;
    std::atomic<unsigned int> cancelIssued{0u};
    // The registered thread's cleanup deadline (CurrentThreadCleanupDeadline of that thread).
    const std::atomic<ULONGLONG>* cleanupUntilTick = nullptr;
};

class Registry final
{
public:
    static Registry& Get() noexcept
    {
        static Registry registry;
        return registry;
    }

    [[nodiscard]] bool Add(Node* node) noexcept
    {
        std::scoped_lock lock(_mutex);
        if (! _timer)
        {
            _timer.reset(CreateThreadpoolTimer(&Registry::TimerCallback, this, nullptr));
            if (! _timer)
            {
                return false;
            }
        }
        // Allocation failure is fatal under the repository policy. Add is noexcept so vector
        // growth intentionally terminates instead of pretending registration was optional.
        _nodes.push_back(node);
        if (! _armed)
        {
            // Relative due time in 100-ns units; negative means "from now".
            FILETIME dueTime{};
            LARGE_INTEGER due{};
            due.QuadPart           = -static_cast<LONGLONG>(kPollIntervalMs) * 10'000;
            dueTime.dwLowDateTime  = due.LowPart;
            dueTime.dwHighDateTime = static_cast<DWORD>(due.HighPart);
            SetThreadpoolTimer(_timer.get(), &dueTime, kPollIntervalMs, 0u);
            _armed = true;
        }
        return true;
    }

    void Remove(Node* node) noexcept
    {
        std::scoped_lock lock(_mutex);
        for (size_t index = 0u; index < _nodes.size(); ++index)
        {
            if (_nodes[index] == node)
            {
                _nodes[index] = _nodes.back();
                _nodes.pop_back();
                break;
            }
        }
        if (_nodes.empty() && _armed && _timer)
        {
            SetThreadpoolTimer(_timer.get(), nullptr, 0u, 0u);
            _armed = false;
        }
    }

    Registry(const Registry&)            = delete;
    Registry& operator=(const Registry&) = delete;

private:
    Registry() = default;

    ~Registry()
    {
        // No scope can be alive here (every scope is stack-scoped inside a worker that the owner
        // joins first), so this only drains a callback that may still be running.
        UniqueThreadpoolTimer timer;
        {
            std::scoped_lock lock(_mutex);
            timer  = std::move(_timer);
            _armed = false;
        }
        if (timer)
        {
            SetThreadpoolTimer(timer.get(), nullptr, 0u, 0u);
            WaitForThreadpoolTimerCallbacks(timer.get(), TRUE);
        }
    }

    static void CALLBACK TimerCallback(PTP_CALLBACK_INSTANCE, void* context, PTP_TIMER) noexcept
    {
        static_cast<Registry*>(context)->Poll();
    }

    void Poll() noexcept
    {
        const ULONGLONG now = GetTickCount64();
        std::scoped_lock lock(_mutex);
        for (Node* node : _nodes)
        {
            if (! node->shouldCancel(node->context))
            {
                node->cancelSinceTick = 0u;
                continue;
            }
            if (node->cancelSinceTick == 0u)
            {
                node->cancelSinceTick = now;
                continue;
            }
            if (now - node->cancelSinceTick < kGraceMs)
            {
                continue;
            }
            if (node->cleanupUntilTick != nullptr && now < node->cleanupUntilTick->load(std::memory_order_acquire))
            {
                // I35 R7.4: owned-stage cleanup has its own budget; it is canceled only once its
                // short control's deadline has passed (a dead share still returns then).
                continue;
            }
            // Re-issued on every poll: a call that starts after an earlier cancel (for example a
            // cleanup delete on the same dead share) must be canceled as well. The count is
            // published before the call so the released thread always observes it.
            node->cancelIssued.fetch_add(1u, std::memory_order_release);
            static_cast<void>(CancelSynchronousIo(node->thread.get()));
        }
    }

    std::mutex _mutex;
    std::vector<Node*> _nodes;
    UniqueThreadpoolTimer _timer;
    bool _armed = false;
};
} // namespace Detail

// Registers the calling thread until destroyed. A null predicate registers nothing (the scope is
// then a no-op, which lets callers construct it unconditionally).
class Scope final
{
public:
    Scope(ShouldCancelFn shouldCancel, void* context) noexcept
    {
        if (shouldCancel == nullptr)
        {
            return;
        }
        if (DuplicateHandle(GetCurrentProcess(),
                            GetCurrentThread(),
                            GetCurrentProcess(),
                            _node.thread.put(),
                            THREAD_TERMINATE, // the access right CancelSynchronousIo requires
                            FALSE,
                            0u) == FALSE)
        {
            return;
        }
        _node.shouldCancel     = shouldCancel;
        _node.context          = context;
        _node.cleanupUntilTick = CurrentThreadCleanupDeadline();
        _registered            = Detail::Registry::Get().Add(&_node);
    }

    ~Scope()
    {
        if (_registered)
        {
            Detail::Registry::Get().Remove(&_node);
        }
    }

    Scope(const Scope&)            = delete;
    Scope& operator=(const Scope&) = delete;
    Scope(Scope&&)                 = delete;
    Scope& operator=(Scope&&)      = delete;

    [[nodiscard]] bool Registered() const noexcept
    {
        return _registered;
    }

    // Number of CancelSynchronousIo calls the watch issued for this scope (diagnostics/tests).
    [[nodiscard]] unsigned int CancelIssuedCount() const noexcept
    {
        return _node.cancelIssued.load(std::memory_order_relaxed);
    }

private:
    Detail::Node _node;
    bool _registered = false;
};

// I35 R7.4: opens the calling thread's owned-stage cleanup window until `deadlineTick` (pass the
// deadline of MakeOwnedStageCleanupOptions; it is capped at kMaxCleanupWindowMs from now) for every
// module's watch, and restores the previous window when destroyed (windows nest). Only the removal or
// restoration of an exact object the operation already owns may run inside one; never primary work.
class CleanupWindow final
{
public:
    explicit CleanupWindow(ULONGLONG deadlineTick) noexcept : _deadline(CurrentThreadCleanupDeadline())
    {
        if (_deadline == nullptr)
        {
            return;
        }
        const ULONGLONG now = GetTickCount64();
        const ULONGLONG cap = now + kMaxCleanupWindowMs;
        _previous           = _deadline->load(std::memory_order_relaxed);
        const ULONGLONG own = deadlineTick < cap ? deadlineTick : cap;
        _effective          = own > _previous ? own : _previous;
        _deadline->store(_effective, std::memory_order_release);
    }

    ~CleanupWindow()
    {
        if (_deadline != nullptr)
        {
            _deadline->store(_previous, std::memory_order_release);
        }
    }

    CleanupWindow(const CleanupWindow&)            = delete;
    CleanupWindow& operator=(const CleanupWindow&) = delete;
    CleanupWindow(CleanupWindow&&)                 = delete;
    CleanupWindow& operator=(CleanupWindow&&)      = delete;

    // The watch reads the window when it polls: a poll that read it just before it opened may cancel a
    // call that then started inside it. A call that failed with ERROR_OPERATION_ABORTED while this
    // window's deadline has not passed was canceled that way, not by its budget; the caller retries it
    // once (the exact-object cleanup calls are idempotent on their own handle).
    [[nodiscard]] bool CanceledInsideWindow(HRESULT hr) const noexcept
    {
        return hr == HRESULT_FROM_WIN32(ERROR_OPERATION_ABORTED) && _effective != 0u && GetTickCount64() < _effective;
    }

private:
    std::atomic<ULONGLONG>* _deadline = nullptr;
    ULONGLONG _previous               = 0u;
    ULONGLONG _effective              = 0u;
};
} // namespace Common::SynchronousIoCancelWatch
