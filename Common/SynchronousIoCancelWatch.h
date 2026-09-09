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

#include <atomic>
#include <mutex>
#include <utility>
#include <vector>

#ifndef WIN32_LEAN_AND_MEAN
#define WIN32_LEAN_AND_MEAN
#endif
#ifndef NOMINMAX
#define NOMINMAX
#endif
#include <Windows.h>

#pragma warning(push)
#pragma warning(disable : 4625 4626 5026 5027 4820 28182) // WIL wrappers are intentionally move-only.
#include <wil/resource.h>
#pragma warning(pop)

namespace Common::SynchronousIoCancelWatch
{
inline constexpr DWORD kPollIntervalMs = 50u;
inline constexpr DWORD kGraceMs        = 500u;

// Returns true when the registered thread's pending synchronous call must be canceled.
using ShouldCancelFn = bool (*)(void* context) noexcept;

namespace Detail
{
using UniqueThreadpoolTimer = wil::unique_any<PTP_TIMER, decltype(&::CloseThreadpoolTimer), ::CloseThreadpoolTimer>;

struct Node final
{
    Node() noexcept = default;
    Node(const Node&)            = delete;
    Node& operator=(const Node&) = delete;
    Node(Node&&)                 = delete;
    Node& operator=(Node&&)      = delete;

    wil::unique_handle thread;
    ShouldCancelFn shouldCancel = nullptr;
    void* context             = nullptr;
    ULONGLONG cancelSinceTick = 0u;
    std::atomic<unsigned int> cancelIssued{0u};
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
            due.QuadPart        = -static_cast<LONGLONG>(kPollIntervalMs) * 10'000;
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
        _node.shouldCancel = shouldCancel;
        _node.context      = context;
        _registered        = Detail::Registry::Get().Add(&_node);
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
} // namespace Common::SynchronousIoCancelWatch
